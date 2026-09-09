import React, { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { createPortal } from 'react-dom';
import { AttachmentQueue, installSendHook } from './queue.js';
import { BASE, DEFAULTS, SENTINEL, READ_EXTS, extOf, imageFile, prettySize } from './protocol.js';
import css from './client.css';
export const name = 'dsh-sendfile/client';
export const inject = ['conversation', 'sessions', 'slots', 'settingsScope'];
const COLORS = { doc: '#3d73b3', docx: '#3d73b3', xls: '#368361', xlsx: '#368361', pptx: '#ba7244', pdf: '#bd6052', md: '#78619b', html: '#497d87' };
const statusText = file => file.status === 'uploading' ? '上传中…' : ['uploaded', 'parsing'].includes(file.status) ? '解析中…' : file.status === 'error' ? file.error || '处理失败，可重试' : `已就绪 · ${(file.chars || file.text?.length || 0).toLocaleString()} 字符`;

export function apply(ctx) {
  const queue = new AttachmentQueue();
  let config = { ...DEFAULTS }; let token; let alive = true;
  const disposers = []; const activeRequests = new Set(); const dialogRoots = new Set(); const history = new Map();
  const style = document.createElement('style'); style.textContent = css; style.id = 'dsh-sendfile-style'; document.head.append(style);
  const currentId = () => ctx.sessions.list.getSnapshot().current;
  const inputFor = sid => { const scope = sid && ctx.sessions.scope(sid); return scope ? ctx.get('conversation').input.for(scope) : null; };
  function notify(message, sid = currentId()) {
    try { inputFor(sid)?.notify?.('error', message); } catch { /* still show the local toast */ }
    document.querySelector('.sf-toast')?.remove();
    const toast = document.createElement('div'); toast.className = 'sf-root sf-toast'; toast.setAttribute('role', 'status'); toast.textContent = message; document.body.append(toast);
    const timer = setTimeout(() => toast.remove(), 7000); disposers.push(() => { clearTimeout(timer); toast.remove(); });
  }
  async function api(action, sid, options = {}, id) {
    const url = `${BASE}/${action}${sid ? `?sessionId=${encodeURIComponent(sid)}` : ''}${id ? `&id=${encodeURIComponent(id)}` : ''}`;
    const controller = new AbortController(); activeRequests.add(controller);
    try {
      const response = await fetch(url, { ...options, signal: controller.signal, headers: { ...options.headers, ...(token ? { 'x-sendfile-token': token } : {}) } });
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.error || '文件服务暂时无法连接，请刷新 DSH 后重试。');
      return result.value;
    } catch (error) { if (error.name === 'SyntaxError') throw new Error('文件服务未加载，请重启 DSH 后重试。'); throw error; }
    finally { activeRequests.delete(controller); }
  }
  const ready = api('config').then(value => { config = value; token = value.token; }).catch(error => { notify(error.message); throw error; });
  // Attach a handler immediately, while retaining the rejecting promise for upload callers.
  ready.catch(() => {});
  function syncSentinel() {
    for (const [sid, items] of queue.items) {
      const input = inputFor(sid); if (!input) continue;
      const draft = input.state.getSnapshot().draft;
      if (items.length && !draft.trim()) input.setDraft(SENTINEL);
      else if (!items.length && draft === SENTINEL) input.setDraft('');
    }
    try {
      const saved = [...queue.items].flatMap(([sid, list]) => list.filter(f => f.id).map(({ id, localId, name, size }) => ({ sessionId: sid, id, localId, name, size })));
      localStorage.setItem('dsh-sendfile.drafts.v1', JSON.stringify(saved));
    } catch { /* Storage quota must not prevent sending. */ }
  }
  disposers.push(queue.subscribe(syncSentinel));
  disposers.push(installSendHook(ctx.get('conversation'), queue, () => config, notify, (sid, ids) => { if (ids.length) api('sent', sid, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids }) }).catch(() => {}); }));
  async function processFile(sid, localId, file, existingId) {
    try {
      await ready;
      if (!existingId && file.size > config.maxFileMB * 1048576) throw new Error(`文件超过 ${config.maxFileMB} MB，请拆分后重发。`);
      const meta = existingId ? { id: existingId } : await api('upload', sid, { method: 'POST', headers: { 'Content-Type': 'application/octet-stream', 'x-sendfile-name': encodeURIComponent(file.name) }, body: file });
      if (!alive || !queue.list(sid).some(f => f.localId === localId)) return;
      queue.update(sid, localId, { ...meta, status: 'parsing' });
      const parsed = await api('parse', sid, { method: 'POST' }, meta.id);
      if (alive) queue.update(sid, localId, { ...parsed, file: null, status: 'ready' });
    } catch (error) { if (alive) queue.update(sid, localId, { status: 'error', error: error.message }); }
  }
  // At most two files upload/parse at once in this window.
  const tasks = []; let working = 0;
  function enqueue(task) { tasks.push(task); drain(); }
  function drain() { while (alive && working < 2 && tasks.length) { working++; Promise.resolve().then(tasks.shift()).finally(() => { working--; drain(); }); } }
  function addFiles(files, sid = currentId()) {
    if (!sid) return notify('请先选择或新建一个带工作区的会话。');
    for (const file of files) {
      if (imageFile(file)) continue;
      if (queue.list(sid).length >= 10) { notify('每条消息最多添加 10 个文档。'); break; }
      const localId = crypto.randomUUID();
      const error = extOf(file.name) === 'ppt' ? '旧版 .ppt 请另存为 .pptx 或 PDF 再发。' : !READ_EXTS.includes(extOf(file.name)) ? '暂不支持此格式。' : null;
      queue.add(sid, { localId, sessionId: sid, name: file.name, size: file.size, file, status: error ? 'error' : 'uploading', error });
      if (!error) enqueue(() => processFile(sid, localId, file));
    }
  }
  function retry(file) {
    if (!file.id && !file.file) return notify('原文件尚未上传，请移除卡片后重新选择文件。');
    queue.update(file.sessionId, file.localId, { status: file.id ? 'parsing' : 'uploading', error: null });
    enqueue(() => processFile(file.sessionId, file.localId, file.file, file.id));
  }
  function useQueue(sid) { const [, redraw] = useState(0); useEffect(() => queue.subscribe(() => redraw(n => n + 1)), []); return queue.list(sid); }
  function Card({ file, removable = false }) {
    return <div className="sf-card"><button className="sf-card-main" onClick={() => preview(file)} disabled={!file.id} title={file.name}>
      <span className="sf-file-icon" style={{ '--sf-color': COLORS[extOf(file.name)] || '#657c63' }}>{extOf(file.name).toUpperCase()}</span>
      <span className="sf-file-label"><span className="sf-filename">{file.name}</span><span className={`sf-status ${file.status === 'error' ? 'sf-error' : ''}`}>{prettySize(file.size)} · {file.status ? statusText(file) : file.truncated ? `已截断 · 已发送 ${file.includedChars.toLocaleString()}/${file.totalChars.toLocaleString()} 字符` : '点击预览'}</span></span>
    </button>{removable && <><span>{file.status === 'error' && <button className="sf-btn sf-icon-button" title="重试" aria-label={`重试 ${file.name}`} onClick={() => retry(file)}>↻</button>}</span><button className="sf-btn sf-icon-button" title="移除附件" aria-label={`移除 ${file.name}`} onClick={() => { queue.remove(file.sessionId, file.localId); if (file.id) api('discard', file.sessionId, { method: 'POST' }, file.id).catch(() => {}); }}>×</button></>}</div>;
  }
  function Dock({ sessionId }) {
    const list = useQueue(sessionId); const count = list.reduce((n, f) => n + (f.chars || 0), 0);
    const [seat, setSeat] = useState(null);
    useEffect(() => {
      let element; let frame;
      const mount = () => {
        frame = null;
        const slot = document.querySelector('[data-composer-card] [data-slot="conversation.input.attachments"]');
        if (!slot || element?.parentElement === slot) return;
        element?.remove(); element = document.createElement('div'); element.style.display = 'contents'; element.dataset.sendfileSeat = 'true'; slot.append(element); setSeat(element);
      };
      mount();
      const observer = new MutationObserver(() => { if (!frame) frame = requestAnimationFrame(mount); });
      observer.observe(document.body, { childList: true, subtree: true });
      return () => { observer.disconnect(); cancelAnimationFrame(frame); element?.remove(); };
    }, []);
    if (!seat || !list.length) return null;
    return createPortal(<div className="sf-root sf-composer-files"><div className="sf-rail">{list.map(f => <Card key={f.localId} file={f} removable />)}</div>{count > config.maxChars && <div className="sf-warning">文件共 {count.toLocaleString()} 字符，发送时将按顺序截断至 {config.maxChars.toLocaleString()} 字符。全文仍可在预览中查看，AI 可按需分段读取。</div>}</div>, seat);
  }
  function Buttons({ sessionId }) {
    const ref = useRef();
    return <div className="sf-root sf-actions"><input ref={ref} type="file" multiple hidden accept={READ_EXTS.map(x => `.${x}`).join(',')} onChange={event => { addFiles([...event.target.files], sessionId); event.target.value = ''; }} />
      <button className="sf-btn sf-attach-button" onClick={() => ref.current.click()} aria-label="添加文档" title="添加文档"><svg className="sf-attach-icon" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true"><path d="M5.5498 9.75V5H6.9502V9.75C6.9502 10.3299 7.4201 10.7998 8 10.7998C8.5799 10.7998 9.0498 10.3299 9.0498 9.75V4.5C9.0498 2.9536 7.7964 1.7002 6.25 1.7002C4.7036 1.7002 3.4502 2.9536 3.4502 4.5V9.75C3.4502 12.2629 5.4871 14.2998 8 14.2998C10.5129 14.2998 12.5498 12.2629 12.5498 9.75V4H13.9502V9.75C13.9502 13.0361 11.2861 15.7002 8 15.7002C4.71391 15.7002 2.0498 13.0361 2.0498 9.75V4.5C2.04981 2.1804 3.9304 0.299806 6.25 0.299805C8.5696 0.299805 10.4502 2.1804 10.4502 4.5V9.75C10.4502 11.1031 9.3531 12.2002 8 12.2002C6.6469 12.2002 5.5498 11.1031 5.5498 9.75Z" fill="currentColor"/></svg></button></div>;
  }
  function Modal({ title, subtitle, children, close, action }) {
    const dialogRef = useRef();
    useEffect(() => {
      const prior = document.activeElement; dialogRef.current?.focus();
      const key = event => {
        if ([...dialogRoots].at(-1) !== close) return;
        if (event.key === 'Escape') { event.stopPropagation(); close(); }
        if (event.key === 'Tab') {
          const controls = [...dialogRef.current.querySelectorAll('button:not(:disabled),a[href],input,iframe')]; const first = controls[0]; const last = controls.at(-1);
          if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
          else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
        }
      };
      window.addEventListener('keydown', key, true); return () => { window.removeEventListener('keydown', key, true); prior?.focus?.(); };
    }, []);
    return <div className="sf-root sf-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) close(); }}><section className="sf-dialog" role="dialog" aria-modal="true" aria-label={title} ref={dialogRef} tabIndex={-1}><header className="sf-dialog-head"><div><h2>{title}</h2><div className="sf-subtitle">{subtitle}</div></div><div className="sf-actions">{action}<button className="sf-btn" onClick={close} aria-label="关闭预览">关闭 ×</button></div></header>{children}</section></div>;
  }
  function mountDialog(renderer) {
    const element = document.createElement('div'); document.body.append(element); const root = createRoot(element);
    const close = () => { root.unmount(); element.remove(); dialogRoots.delete(close); };
    dialogRoots.add(close); root.render(renderer(close)); return close;
  }
  function Preview({ initial, close }) {
    const [file, setFile] = useState(initial); const [error, setError] = useState(''); const [mode, setMode] = useState(['pdf', 'html'].includes(extOf(initial.name)) ? 'page' : 'text');
    useEffect(() => { let active = true; api('document', initial.sessionId, {}, initial.id).then(value => { if (active) setFile(value); }).catch(e => { if (active) setError(e.message); }); return () => { active = false; }; }, []);
    const url = `${BASE}/file?sessionId=${encodeURIComponent(file.sessionId)}&id=${encodeURIComponent(file.id)}`;
    return <Modal title={file.name} subtitle={`${prettySize(file.size)} · ${file.kind === 'output' ? '生成的文件' : '上传的文件'} · 原件保持不动`} close={close} action={<a className="sf-btn sf-primary" href={url} download={file.name}>下载</a>}><div className="sf-dialog-body">
      {['pdf', 'html'].includes(extOf(file.name)) && <div className="sf-tabs"><button className="sf-btn" onClick={() => setMode('page')}>页面预览</button><button className="sf-btn" onClick={() => setMode('text')}>提取内容</button></div>}
      {file.warning && <p className="sf-warning">{file.warning}</p>}{error ? <p className="sf-error">{error}</p> : mode === 'page' ? <iframe title={file.name} className="sf-iframe" sandbox="allow-scripts" src={`${url}&inline=1`} /> : <><p className="sf-note">内容预览用于核对文字和数据，Office 的原始排版、图表及嵌入图片请下载后核对。</p><pre className="sf-document">{file.text ?? file.error ?? '正在读取…'}</pre></>}
    </div></Modal>;
  }
  function preview(file) { mountDialog(close => <Preview initial={file} close={close} />); }
  function SettingsRow() {
    const [maxChars, setMaxChars] = useState(config.maxChars); const [maxFileMB, setMaxFileMB] = useState(config.maxFileMB); const [retentionDays, setRetentionDays] = useState(config.retentionDays ?? 7); const [message, setMessage] = useState('');
    async function save() { try { await ready; const value = await api('config', null, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ maxChars: Number(maxChars), maxFileMB: Number(maxFileMB), retentionDays: Number(retentionDays) }) }); config = value; token = value.token; queue.emit(); setMessage('设置已保存。'); } catch (e) { setMessage(e.message); } }
    return <section className="sf-root sf-settings-row"><div className="sf-settings-heading"><strong>发文件插件</strong><span>文档在本机解析，发送后文字交给当前模型</span></div><div className="sf-settings-fields"><label>每条消息最多发送的文档字符数<input type="number" min="1000" max="200000" value={maxChars} onChange={e => setMaxChars(e.target.value)} /></label><p className="sf-note">模型上下文还包括历史消息和工具。出现上下文超限时请调低。</p><label>单文件大小上限（MB）<input type="number" min="1" max="64" value={maxFileMB} onChange={e => setMaxFileMB(e.target.value)} /></label><label>发送成功后副本保留天数（0 为不自动清理）<input type="number" min="0" max="90" value={retentionDays} onChange={e => setRetentionDays(e.target.value)} /></label><div className="sf-settings-save"><button className="sf-btn sf-primary" onClick={save}>保存设置</button><span role="status">{message}</span></div><p className="sf-note">不调用云端解析或 OCR，不运行文件中的宏。扫描 PDF 与旧版 PPT 请先在本机转换。</p></div></section>;
  }
  for (const [slot, component, order] of [['conversation.input.left', Buttons, 5], ['conversation.input.dock', Dock, 5]]) {
    disposers.push(ctx.slots.inject(slot, () => ctx.slots.register({ name: slot, id: 'sendfile', order }, component)));
  }
  disposers.push(ctx.slots.inject('settings.general.item', () => ctx.slots.register({ name: 'settings.general.item', id: 'sendfile-settings', order: 90 }, SettingsRow)));
  const filePayload = event => [...(event.dataTransfer?.types || [])].includes('Files');
  const over = event => { if (filePayload(event)) { event.preventDefault(); event.dataTransfer.dropEffect = 'copy'; } };
  const drop = event => {
    if (!filePayload(event)) return;
    const files = [...event.dataTransfer.files]; const docs = files.filter(f => !imageFile(f)); if (!docs.length) return;
    event.preventDefault(); event.stopPropagation(); addFiles(docs);
    const images = files.filter(imageFile);
    if (images.length) { const data = new DataTransfer(); images.forEach(f => data.items.add(f)); event.target.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: data })); }
    window.dispatchEvent(new DragEvent('dragend'));
  };
  for (const [type, handler, capture] of [['dragover', over, false], ['drop', drop, true]]) { window.addEventListener(type, handler, capture); disposers.push(() => window.removeEventListener(type, handler, capture)); }

  // Version-specific presentation adapter. Original DSH-owned nodes stay intact;
  // the complete prompt remains in the message record and native copy action.
  function reconcileHistory() {
    for (const [original, view] of history) if (!original.isConnected || !original.textContent.includes('[[SENDFILE-META ')) { original.classList.remove('sf-original'); view.root.unmount(); view.element.remove(); history.delete(original); }
    for (const original of document.querySelectorAll('[data-time-hover-root] [class*="_bubble"]')) {
      if (history.has(original) || original.closest('.sf-history')) continue;
      const text = original.textContent || ''; const match = text.match(/\[\[SENDFILE-META (.+)\]\]\n/);
      if (!match) continue;
      try {
        const files = JSON.parse(match[1]);
        if (!Array.isArray(files) || !files.length || !files.every(f => typeof f.id === 'string' && typeof f.sessionId === 'string')) continue;
        const element = document.createElement('div'); element.className = 'sf-root sf-history'; original.after(element);
        const root = createRoot(element); root.render(<><div className="sf-history-prompt">{text.slice(0, match.index).trim()}</div><div className="sf-rail">{files.map(file => <Card file={file} key={file.id} />)}</div><span className="sf-preview-label">文档文字已随消息发送</span></>);
        original.classList.add('sf-original'); history.set(original, { element, root });
      } catch { /* A user-authored lookalike marker must keep its normal rendering. */ }
    }
  }
  let historyTimer;
  const observer = new MutationObserver(() => { if (!historyTimer) historyTimer = setTimeout(() => { historyTimer = null; reconcileHistory(); }, 100); });
  observer.observe(document.body, { childList: true, characterData: true, subtree: true }); reconcileHistory();
  const click = event => {
    const anchor = event.target.closest?.('a[href]'); if (!anchor) return;
    let url; try { url = new URL(anchor.href, location.href); } catch { return; }
    if (url.origin !== location.origin || url.pathname !== `${BASE}/view`) return;
    event.preventDefault(); event.stopPropagation();
    const sid = url.searchParams.get('sessionId'); const id = url.searchParams.get('id');
    api('document', sid, {}, id).then(preview).catch(e => notify(e.message));
  };
  document.addEventListener('click', click, true);
  // Restore uploaded drafts using file IDs only; document text stays on disk.
  try {
    const saved = JSON.parse(localStorage.getItem('dsh-sendfile.drafts.v1') || '[]').slice(0, 50);
    for (const meta of saved) { queue.add(meta.sessionId, { ...meta, status: 'parsing' }); enqueue(() => processFile(meta.sessionId, meta.localId, null, meta.id)); }
  } catch { /* Corrupt local drafts do not block new uploads. */ }
  ctx.effect(() => () => {
    alive = false; tasks.length = 0; observer.disconnect(); clearTimeout(historyTimer); hideOverlay();
    for (const controller of activeRequests) controller.abort();
    document.removeEventListener('click', click, true);
    for (const close of [...dialogRoots]) close();
    for (const [original, view] of history) { original.classList.remove('sf-original'); view.root.unmount(); view.element.remove(); }
    for (const dispose of disposers.reverse()) dispose?.(); style.remove();
  }, 'sendfile client lifecycle');
}
