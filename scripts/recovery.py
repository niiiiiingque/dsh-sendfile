#!/usr/bin/env python3
"""Local DSH snapshot/recovery. Never uploads content or deletes projects."""
import argparse, hashlib, json, os, shutil, stat, subprocess, tempfile
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_HOME = Path.home() / '.dsh'
SNAPSHOT_ITEMS = ['profiles/desktop', 'sessions', 'storages', 'attachments', 'settings.yaml']

def stamp():
    return datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')

def assert_stopped():
    result = subprocess.run(['/usr/bin/pgrep', '-f', '/Applications/DSH Desktop.app/Contents/'], capture_output=True)
    if result.returncode == 0:
        raise RuntimeError('DSH Desktop 仍在运行；请先退出。')
    if result.returncode != 1:
        raise RuntimeError('无法核对 DSH 进程，未执行变更。')

def sha(file):
    value = hashlib.sha256()
    with file.open('rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            value.update(chunk)
    return value.hexdigest()

def inventory(root):
    result = {}
    def visit(file, relative):
        mode = file.lstat().st_mode
        entry = {'mode': stat.S_IMODE(mode)}
        if stat.S_ISLNK(mode):
            entry.update(type='symlink', target=os.readlink(file))
        elif stat.S_ISDIR(mode):
            entry['type'] = 'directory'
            for child in sorted(file.iterdir()): visit(child, relative + '/' + child.name)
        elif stat.S_ISREG(mode):
            entry.update(type='file', size=file.stat().st_size, sha256=sha(file))
        else:
            raise RuntimeError('备份范围出现非普通文件，未继续：' + relative)
        result[relative] = entry
    visit(root, '.')
    return result

def copy_item(source, target):
    target.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    if source.is_symlink(): os.symlink(os.readlink(source), target)
    elif source.is_dir(): shutil.copytree(source, target, symlinks=True, copy_function=shutil.copy2)
    else: shutil.copy2(source, target)

def backup(dsh_home):
    destination = dsh_home / 'sendfile-backups' / stamp()
    destination.mkdir(parents=True, mode=0o700)
    os.chmod(destination.parent, 0o700)
    manifest = {'format': 1, 'dshHome': str(dsh_home.resolve()), 'createdAt': stamp(), 'items': {}}
    for name in SNAPSHOT_ITEMS:
        source = dsh_home / name
        if not source.exists() and not source.is_symlink(): continue
        print('备份并校验：' + name, flush=True)
        before = inventory(source)
        copy_item(source, destination / 'snapshot' / name)
        after = inventory(destination / 'snapshot' / name)
        if before != after: raise RuntimeError('备份内容不一致，未安装：' + name)
        if inventory(source) != before: raise RuntimeError('备份期间源文件发生变化，未安装：' + name)
        manifest['items'][name] = before
    (destination / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2))
    os.chmod(destination / 'manifest.json', 0o600)
    print(json.dumps({'backup': str(destination), 'verified': True,
        'files': sum(e['type'] == 'file' for item in manifest['items'].values() for e in item.values()),
        'bytes': sum(e.get('size', 0) for item in manifest['items'].values() for e in item.values())}, ensure_ascii=False), flush=True)
    return destination

def verify(dsh_home, folder):
    folder = folder.resolve()
    if folder.parent != (dsh_home / 'sendfile-backups').resolve():
        raise RuntimeError('仅接受此 DSH_HOME 的直接备份子目录。')
    manifest = json.loads((folder / 'manifest.json').read_text())
    if manifest.get('format') != 1 or manifest.get('dshHome') != str(dsh_home.resolve()):
        raise RuntimeError('备份归属或格式不匹配。')
    if 'profiles/desktop' not in manifest['items']: raise RuntimeError('备份缺少 desktop profile。')
    for name, entries in manifest['items'].items():
        if name not in SNAPSHOT_ITEMS: raise RuntimeError('备份项目无效。')
        if inventory(folder / 'snapshot' / name) != entries: raise RuntimeError('备份校验失败：' + name)
    return manifest

def restore_profile(dsh_home, folder):
    manifest = verify(dsh_home, folder)
    profiles = dsh_home / 'profiles'
    staging = profiles / ('.desktop-restore-' + stamp())
    retained = profiles / ('.desktop-before-restore-' + stamp())
    source = folder / 'snapshot/profiles/desktop'
    copy_item(source, staging)
    if inventory(staging) != manifest['items']['profiles/desktop']:
        raise RuntimeError('恢复副本校验失败，当前环境未改变。')
    current = profiles / 'desktop'
    moved = False
    if current.exists(): current.rename(retained); moved = True
    try: staging.rename(current)
    except BaseException:
        if moved: retained.rename(current)
        raise
    if inventory(current) != manifest['items']['profiles/desktop']:
        raise RuntimeError('恢复后校验失败，请保持 DSH 退出状态。')
    print(json.dumps({'restored': str(current), 'previousEnvironmentRetained': str(retained) if moved else None,
        'sessionAndProjectFilesUntouched': True}, ensure_ascii=False), flush=True)

def selftest():
    with tempfile.TemporaryDirectory(prefix='dsh-recovery-test-') as tmp:
        root = Path(tmp)
        profile = root / 'profiles/desktop'
        (profile / 'node_modules/library').mkdir(parents=True)
        (profile / 'package.json').write_text('{"before":true}')
        (profile / 'node_modules/library/index.js').write_text('original dependency')
        os.symlink('library', profile / 'node_modules/alias')
        (root / 'sessions').mkdir()
        (root / 'sessions/original').write_text('original history')
        saved = backup(root)
        (profile / 'package.json').write_text('{"broken":true}')
        (profile / 'node_modules/library/index.js').write_text('changed dependency')
        (root / 'sessions/new').write_text('new history must survive recovery')
        restore_profile(root, saved)
        assert (profile / 'package.json').read_text() == '{"before":true}'
        assert (profile / 'node_modules/library/index.js').read_text() == 'original dependency'
        assert os.readlink(profile / 'node_modules/alias') == 'library'
        assert (root / 'sessions/new').exists()
        (saved / 'snapshot/profiles/desktop/package.json').write_text('tampered')
        try: verify(root, saved)
        except RuntimeError: pass
        else: raise AssertionError('Tampered backup accepted')
    print('SELFTEST PASS: restores dependencies and symlinks; preserves newer sessions; rejects damaged backup.')

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['backup', 'verify', 'restore-profile', 'selftest'])
    parser.add_argument('folder', nargs='?', type=Path)
    parser.add_argument('--apply', action='store_true')
    args = parser.parse_args()
    if args.action == 'selftest': selftest()
    elif args.action == 'verify':
        if args.folder is None: parser.error('需要备份目录')
        verify(DEFAULT_HOME, args.folder); print('备份逐文件校验通过。')
    elif not args.apply: print('仅显示计划。实际执行需 --apply；restore-profile 只恢复插件环境，保留当前会话和项目文件。')
    else:
        assert_stopped()
        if args.action == 'backup': backup(DEFAULT_HOME)
        elif args.folder is None: parser.error('需要备份目录')
        else: restore_profile(DEFAULT_HOME, args.folder.resolve())
