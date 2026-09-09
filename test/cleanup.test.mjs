import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { shouldPurge, FileStore } from '../src/storage.js';
import { clampSettings, DEFAULTS } from '../src/protocol.js';

test('shouldPurge 只清上传副本，不动生成文件', () => {
  const now = Date.parse('2026-09-09T12:00:00Z');
  assert.equal(shouldPurge({ kind: 'output', createdAt: '2020-01-01T00:00:00Z' }, now, 7), false);
  assert.equal(shouldPurge({ kind: 'input', createdAt: new Date(now - 31 * 86400000).toISOString() }, now, 7), true);
  assert.equal(shouldPurge({ kind: 'input', createdAt: new Date(now - 29 * 86400000).toISOString() }, now, 7), false);
  assert.equal(shouldPurge({ kind: 'input', createdAt: '2026-09-01T00:00:00Z', sentAt: new Date(now - 8 * 86400000).toISOString() }, now, 7), true);
  assert.equal(shouldPurge({ kind: 'input', createdAt: '2026-09-01T00:00:00Z', sentAt: new Date(now - 6 * 86400000).toISOString() }, now, 7), false);
  assert.equal(shouldPurge({ kind: 'input', createdAt: '2026-09-01T00:00:00Z', sentAt: new Date(now - 400 * 86400000).toISOString() }, now, 0), false);
  assert.equal(shouldPurge({ kind: 'input', createdAt: '坏日期' }, now, 7), false);
});
test('clampSettings 接受 retentionDays 并限幅', () => {
  assert.equal(clampSettings({}).retentionDays, DEFAULTS.retentionDays);
  assert.equal(clampSettings({ retentionDays: 3 }).retentionDays, 3);
  assert.equal(clampSettings({ retentionDays: 0 }).retentionDays, 0);
  assert.throws(() => clampSettings({ retentionDays: 91 }));
  assert.throws(() => clampSettings({ retentionDays: 1.5 }));
});
test('sweepAll 清理过期上传副本并保留生成文件与新副本', async () => {
  const ws = await mkdtemp(path.join(tmpdir(), 'sendfile-sweep-'));
  const sessionId = 'sweep-session';
  const store = new FileStore({ list: () => [{ path: ws, sessionIds: [sessionId] }] });
  const old = await store.put(sessionId, '旧上传.txt', Buffer.from('x'), 'input');
  const fresh = await store.put(sessionId, '新上传.txt', Buffer.from('y'), 'input');
  const output = await store.put(sessionId, '交付.docx', Buffer.from('z'), 'output');
  await store.markSent(sessionId, old.id);
  const metaPath = path.join(ws, '.dsh-sendfile', createHash('sha256').update(sessionId).digest('hex').slice(0, 24), old.id, 'meta.json');
  const meta = JSON.parse(await readFile(metaPath, 'utf8'));
  meta.sentAt = new Date(Date.now() - 8 * 86400000).toISOString();
  await writeFile(metaPath, JSON.stringify(meta));
  assert.equal(await store.sweepAll([ws], 7), 1);
  await assert.rejects(() => store.get(sessionId, old.id));
  assert.ok(await store.get(sessionId, fresh.id));
  assert.ok(await store.get(sessionId, output.id));
  await rm(ws, { recursive: true, force: true });
});
