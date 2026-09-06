import { randomBytes, timingSafeEqual } from 'node:crypto';
import { BASE, READ_EXTS, extOf, clampSettings } from './protocol.js';
export const ROUTES = ['config', 'upload', 'parse', 'list', 'document', 'file', 'view'];
const escape = text => String(text).replace(/[&<>"']/g, x => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[x]);
export const fileURL = (meta, action = 'view') => `${BASE}/${action}?sessionId=${encodeURIComponent(meta.sessionId)}&id=${encodeURIComponent(meta.id)}`;
export const publicMeta = (meta, origin = '') => ({ ...meta, previewUrl: origin + fileURL(meta), downloadUrl: origin + fileURL(meta, 'file') });
export function sendJSON(res, status, value) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  res.end(JSON.stringify(value));
}
async function readBody(req, max) {
  const chunks = []; let size = 0;
  if (Number(req.headers['content-length']) > max) throw new Error('文件体积超过当前设置的上限。');
  for await (const chunk of req) { size += chunk.length; if (size > max) throw new Error('文件体积超过当前设置的上限。'); chunks.push(chunk); }
  return Buffer.concat(chunks);
}
export function sameOrigin(req) {
  if (['cross-site', 'same-site'].includes(req.headers['sec-fetch-site'])) return false;
  const origin = req.headers.origin;
  if (!origin) return true;
  try { return new URL(origin).host === req.headers.host; } catch { return false; }
}
export function createRouter(store, { getSettings, setSettings }) {
  const token = randomBytes(32).toString('hex');
  let inFlight = 0;
  function authorized(req) { const value = req.headers['x-sendfile-token']; return typeof value === 'string' && /^[a-f0-9]{64}$/.test(value) && timingSafeEqual(Buffer.from(value), Buffer.from(token)); }
  return async (req, res) => {
    let busy = false;
    try {
      const peer = req.socket?.remoteAddress;
      const authority = new URL(`http://${req.headers.host || ''}`).hostname;
      if (!['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(peer) || !['localhost', '127.0.0.1', '[::1]'].includes(authority)) return sendJSON(res, 403, { ok: false, error: '当前版本只支持本机 DSH 窗口。' });
      if (!sameOrigin(req)) return sendJSON(res, 403, { ok: false, error: '请从 DSH 当前窗口访问文件。' });
      const url = new URL(req.url, 'http://localhost'); const action = url.pathname.slice(BASE.length + 1);
      if (!ROUTES.includes(action)) return sendJSON(res, 404, { ok: false, error: '接口不存在。' });
      const method = req.method;
      if (!['GET', 'POST'].includes(method)) return sendJSON(res, 405, { ok: false, error: '不支持此请求方法。' });
      const mutable = ['upload', 'parse'];
      if ((mutable.includes(action) && method !== 'POST') || (!mutable.includes(action) && action !== 'config' && method !== 'GET')) return sendJSON(res, 405, { ok: false, error: '请求方法错误。' });
      if (method === 'POST' && !authorized(req)) return sendJSON(res, 403, { ok: false, error: '连接已更新，请刷新 DSH 后重试。' });
      if (action === 'config') {
        if (method === 'POST') await setSettings(clampSettings(JSON.parse((await readBody(req, 4096)).toString('utf8'))));
        return sendJSON(res, 200, { ok: true, value: { ...getSettings(), token, version: '0.1.2' } });
      }
      const sessionId = url.searchParams.get('sessionId'); const id = url.searchParams.get('id');
      store.workspace(sessionId);
      if (action === 'upload' || action === 'parse') {
        if (inFlight >= 4) return sendJSON(res, 429, { ok: false, error: '正在处理其他文件，请稍后重试。' });
        inFlight++; busy = true;
      }
      if (action === 'upload') {
        const name = decodeURIComponent(req.headers['x-sendfile-name'] || '');
        if (extOf(name) === 'ppt') throw new Error('暂不支持旧版 .ppt，请另存为 .pptx 或 PDF 再发。');
        if (!READ_EXTS.includes(extOf(name))) throw new Error('不支持这种附件格式；图片请使用 DSH 原生图片入口。');
        const bytes = await readBody(req, getSettings().maxFileMB * 1048576);
        if (!bytes.length) throw new Error('文件为空，请检查后重试。');
        const meta = await store.put(sessionId, name, bytes);
        return sendJSON(res, 200, { ok: true, value: publicMeta(meta) });
      }
      if (action === 'parse') return sendJSON(res, 200, { ok: true, value: publicMeta(await store.parse(sessionId, id)) });
      if (action === 'list') return sendJSON(res, 200, { ok: true, value: (await store.list(sessionId)).map(meta => publicMeta(meta)) });
      if (action === 'document') return sendJSON(res, 200, { ok: true, value: publicMeta(await store.get(sessionId, id, true)) });
      const meta = await store.get(sessionId, id, action === 'view');
      if (action === 'file') {
        const bytes = await store.bytes(sessionId, id);
        const inline = url.searchParams.get('inline') === '1';
        const type = inline && meta.format === 'pdf' ? 'application/pdf' : inline && meta.format === 'html' ? 'text/html; charset=utf-8' : 'application/octet-stream';
        res.writeHead(200, {
          'Content-Type': type, 'Content-Length': bytes.length, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff',
          'Content-Disposition': `${inline && ['pdf', 'html'].includes(meta.format) ? 'inline' : 'attachment'}; filename="download.${meta.format}"; filename*=UTF-8''${encodeURIComponent(meta.name).replace(/'/g, '%27')}`,
          'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'; img-src data:; script-src 'unsafe-inline'; sandbox allow-scripts; base-uri 'none'; form-action 'none'",
        }); res.end(bytes); return;
      }
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store', 'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'; frame-ancestors 'self'; base-uri 'none'; form-action 'none'", 'X-Content-Type-Options': 'nosniff' });
      res.end(`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>${escape(meta.name)}</title><style>body{max-width:1000px;margin:40px auto;padding:0 24px;font:16px/1.7 system-ui;color:#253c34;background:#f8f8f4}h1{font-size:24px}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:inherit;padding:24px;background:white;border:1px solid #e1e5df;border-radius:16px}a{color:#26765d}</style><h1>${escape(meta.name)}</h1><p>内容预览 · <a href="${escape(fileURL(meta, 'file'))}" download>下载文件</a></p><pre>${escape(meta.text ?? meta.error ?? '文件尚未完成解析。')}</pre></html>`);
    } catch (error) {
      if (!res.headersSent) sendJSON(res, 400, { ok: false, error: /ENOENT/.test(error.message) ? '文件已不在当前工作区，或尚未同步完成。' : /EACCES|EPERM|\/Users\//.test(error.message) ? '当前工作区无法读写此文件，请检查文件权限。' : error.message });
      else res.end();
    } finally { if (busy) inFlight--; }
  };
}
