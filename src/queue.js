import { SENTINEL, composeMessage } from './protocol.js';
export class AttachmentQueue {
  constructor() { this.items = new Map(); this.listeners = new Set(); this.flights = new Map(); }
  subscribe = listener => { this.listeners.add(listener); return () => this.listeners.delete(listener); };
  emit() { for (const listener of this.listeners) listener(); }
  list(sessionId) { return this.items.get(sessionId) || []; }
  add(sessionId, item) { this.items.set(sessionId, [...this.list(sessionId), item]); this.emit(); }
  update(sessionId, localId, patch) { this.items.set(sessionId, this.list(sessionId).map(item => item.localId === localId ? { ...item, ...patch } : item)); this.emit(); }
  remove(sessionId, localId) { this.items.set(sessionId, this.list(sessionId).filter(item => item.localId !== localId)); this.emit(); }
  take(sessionId) { const batch = this.list(sessionId); this.items.set(sessionId, []); this.emit(); return batch; }
  restore(sessionId, batch) { this.items.set(sessionId, [...batch, ...this.list(sessionId)]); this.emit(); }
}
export function installSendHook(conversation, queue, settings, notify) {
  const proto = Object.getPrototypeOf(conversation);
  if (!proto || typeof proto.sendSession !== 'function') throw new Error('此 DSH 版本的发送接口不兼容。');
  const original = proto.sendSession;
  if (original.__sendfile) throw new Error('发文件插件重复加载，请刷新 DSH。');
  async function wrapped(session, text, imageIds, mode, signal) {
    const sessionId = session.sessionId;
    const pending = queue.list(sessionId);
    if (!pending.length) return original.call(this, session, String(text || '').replaceAll(SENTINEL, ''), imageIds, mode, signal);
    if (pending.some(item => item.status !== 'ready')) {
      notify('附件尚未准备好，请等待解析完成，或移除失败文件后再发送。', sessionId);
      return { kind: 'error' };
    }
    const batch = queue.take(sessionId);
    let succeeded = false;
    try {
      const result = await original.call(this, session, composeMessage(text, batch, settings().maxChars), imageIds, mode, signal);
      succeeded = result?.kind === 'success'; return result;
    } finally { if (!succeeded) queue.restore(sessionId, batch); }
  }
  wrapped.__sendfile = true; proto.sendSession = wrapped;
  return () => { if (proto.sendSession === wrapped) proto.sendSession = original; };
}
