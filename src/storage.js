import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { parseBytes } from './parser.js';
import { extOf } from './protocol.js';

export function safeName(name) {
  const value = String(name).normalize('NFC').replace(/[\\/\u0000-\u001f\u007f<>:"|?*]/g, '_').replace(/^\.+/, '').trim();
  if (!value || Buffer.byteLength(value) > 220) throw new Error('文件名为空或过长，请改名后重试。');
  return value;
}
const idPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
export function contained(root, target) { const relative = path.relative(root, target); return relative !== '..' && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative); }

export function shouldPurge(meta, nowMs, retentionDays) {
  if (!meta || meta.kind !== 'input') return false;
  const created = Date.parse(meta.createdAt);
  if (Number.isNaN(created)) return false;
  if (meta.sentAt) {
    const sent = Date.parse(meta.sentAt);
    if (Number.isNaN(sent) || retentionDays <= 0) return false;
    return nowMs - sent > retentionDays * 86400000;
  }
  return nowMs - created > 30 * 86400000;
}
export class FileStore {
  constructor(registry) { this.registry = registry; this.pending = new Map(); }
  workspace(sessionId) {
    if (typeof sessionId !== 'string' || sessionId.length > 200) throw new Error('无效的会话。');
    const workspace = this.registry.list().find(w => w.sessionIds?.includes(sessionId));
    if (!workspace) throw new Error('当前会话尚未关联工作区，请先在 DSH 选择工作区。');
    return workspace.path;
  }
  async root(sessionId, create = false) {
    const workspace = await fs.realpath(this.workspace(sessionId));
    const hash = createHash('sha256').update(sessionId).digest('hex').slice(0, 24);
    let dir = workspace;
    for (const part of ['.dsh-sendfile', hash]) {
      dir = path.join(dir, part);
      if (create) await fs.mkdir(dir, { mode: 0o700 }).catch(e => { if (e.code !== 'EEXIST') throw e; });
      const stat = await fs.lstat(dir);
      if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error('附件目录不应为符号链接或普通文件。');
      if (!contained(workspace, await fs.realpath(dir))) throw new Error('附件目录超出当前工作区。');
    }
    return dir;
  }
  async entry(sessionId, id) {
    if (!idPattern.test(id)) throw new Error('无效的附件标识。');
    const root = await this.root(sessionId);
    const dir = path.join(root, id);
    const stat = await fs.lstat(dir);
    if (!stat.isDirectory() || stat.isSymbolicLink() || !contained(root, await fs.realpath(dir))) throw new Error('附件目录无效。');
    return dir;
  }
  async regularFile(dir, name) {
    const file = path.join(dir, name);
    const stat = await fs.lstat(file);
    if (!stat.isFile() || stat.isSymbolicLink() || !contained(dir, await fs.realpath(file))) throw new Error('附件文件无效。');
    return file;
  }
  async get(sessionId, id, withText = false) {
    const dir = await this.entry(sessionId, id);
    const metaFile = await this.regularFile(dir, 'meta.json');
    const meta = JSON.parse(await fs.readFile(metaFile, 'utf8'));
    if (meta.sessionId !== sessionId || meta.id !== id || safeName(meta.name) !== meta.name) throw new Error('附件与当前会话不匹配。');
    const value = { ...meta, path: path.join(dir, 'original', meta.name) };
    if (withText && meta.status === 'ready') value.text = await fs.readFile(await this.regularFile(dir, 'content.txt'), 'utf8');
    return value;
  }
  async list(sessionId) {
    let root;
    try { root = await this.root(sessionId); } catch (e) { if (e.code === 'ENOENT') return []; throw e; }
    const ids = (await fs.readdir(root)).filter(x => idPattern.test(x));
    const values = await Promise.all(ids.slice(-1000).map(id => this.get(sessionId, id).catch(() => null)));
    return values.filter(Boolean).sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }
  async update(sessionId, id, change) {
    const dir = await this.entry(sessionId, id);
    const { path: ignoredPath, ...previous } = await this.get(sessionId, id);
    const next = { ...previous, ...change };
    const temp = path.join(dir, `meta-${randomUUID()}.tmp`);
    await fs.writeFile(temp, JSON.stringify(next, null, 2), { flag: 'wx', mode: 0o600 });
    await fs.rename(temp, path.join(dir, 'meta.json'));
    return next;
  }
  async put(sessionId, name, bytes, kind = 'input', provenance = {}) {
    name = safeName(name);
    const root = await this.root(sessionId, true);
    const id = randomUUID(); const dir = path.join(root, id);
    await fs.mkdir(dir, { mode: 0o700 });
    await fs.mkdir(path.join(dir, 'original'), { mode: 0o700 });
    await fs.writeFile(path.join(dir, 'original', name), bytes, { flag: 'wx', mode: 0o600 });
    const meta = { id, sessionId, name, size: bytes.length, format: extOf(name), kind, status: 'uploaded', createdAt: new Date().toISOString(), ...provenance };
    await fs.writeFile(path.join(dir, 'meta.json'), JSON.stringify(meta, null, 2), { flag: 'wx', mode: 0o600 });
    return { ...meta, path: path.join(dir, 'original', name) };
  }
  async bytes(sessionId, id) {
    const meta = await this.get(sessionId, id);
    const dir = await this.entry(sessionId, id); const original = path.join(dir, 'original');
    const stat = await fs.lstat(original);
    if (!stat.isDirectory() || stat.isSymbolicLink() || !contained(dir, await fs.realpath(original))) throw new Error('附件源文件目录无效。');
    return fs.readFile(await this.regularFile(original, meta.name));
  }
  async parse(sessionId, id) {
    const key = `${sessionId}/${id}`;
    if (this.pending.has(key)) return this.pending.get(key);
    const promise = this.doParse(sessionId, id).finally(() => this.pending.delete(key));
    this.pending.set(key, promise); return promise;
  }
  async doParse(sessionId, id) {
    const meta = await this.get(sessionId, id);
    if (meta.status === 'ready') return this.get(sessionId, id, true);
    await this.update(sessionId, id, { status: 'parsing', error: null });
    try {
      const text = await parseBytes(await this.bytes(sessionId, id), meta.name);
      const dir = await this.entry(sessionId, id);
      const temp = path.join(dir, `content-${randomUUID()}.tmp`);
      await fs.writeFile(temp, text, { flag: 'wx', mode: 0o600 });
      await fs.rename(temp, path.join(dir, 'content.txt'));
      await this.update(sessionId, id, { status: 'ready', chars: text.length, error: null });
      return this.get(sessionId, id, true);
    } catch (error) {
      await this.update(sessionId, id, { status: 'error', error: error.message });
      throw error;
    }
  }
  async markSent(sessionId, id) {
    const current = await this.get(sessionId, id);
    if (current.kind !== 'input' || current.sentAt) return current;
    return this.update(sessionId, id, { sentAt: new Date().toISOString() });
  }
  async remove(sessionId, id) {
    const dir = await this.entry(sessionId, id);
    const meta = await this.get(sessionId, id);
    if (meta.kind !== 'input') throw new Error('生成的文件不会被自动清理，请在工作区中自行管理。');
    await fs.rm(dir, { recursive: true, force: true });
    return true;
  }
  async sweepAll(workspacePaths, retentionDays) {
    const now = Date.now();
    let purged = 0;
    for (const ws of workspacePaths) {
      let real;
      try { real = await fs.realpath(String(ws)); } catch { continue; }
      const root = path.join(real, '.dsh-sendfile');
      let hashDirs;
      try { const st = await fs.lstat(root); if (!st.isDirectory() || st.isSymbolicLink()) continue; hashDirs = await fs.readdir(root); } catch { continue; }
      for (const hash of hashDirs) {
        if (!/^[0-9a-f]{24}$/.test(hash)) continue;
        const hashDir = path.join(root, hash);
        let ids;
        try { ids = await fs.readdir(hashDir); } catch { continue; }
        for (const id of ids) {
          if (!idPattern.test(id)) continue;
          const dir = path.join(hashDir, id);
          try {
            const st = await fs.lstat(dir);
            if (!st.isDirectory() || st.isSymbolicLink()) continue;
            const meta = JSON.parse(await fs.readFile(path.join(dir, 'meta.json'), 'utf8'));
            if (shouldPurge(meta, now, retentionDays) && contained(root, await fs.realpath(dir))) { await fs.rm(dir, { recursive: true, force: true }); purged++; }
          } catch { /* 读不动的条目保持原样，不误删 */ }
        }
      }
    }
    return purged;
  }
}
