import { fork, spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { extOf, READ_EXTS } from './protocol.js';
const hints = {
  unsupported: '未能提取文字。若为扫描 PDF，请先在本机 OCR 后再上传。',
  malformed: '文件结构损坏或格式与扩展名不符，请尝试另存后重发。',
  encrypted: '文件有密码或加密保护，请先解除保护再上传。',
  resourceLimit: '文档结构或解压体积过大，请拆分后重发。',
  missingPart: '文档缺少必要部件，请用 Office 或 WPS 另存后重发。',
};
export async function parseBytes(bytes, name, { timeoutMs = 45000 } = {}) {
  const format = extOf(name);
  if (format === 'ppt') throw new Error('暂不支持旧版 .ppt，请另存为 .pptx 或 PDF 再发。');
  if (!READ_EXTS.includes(format) && format !== 'html') throw new Error('暂不支持这种格式。支持 Word、Excel、PPTX、文字 PDF、Markdown、TXT、CSV。');
  if (!bytes.length) throw new Error('文件为空，请检查后重试。');
  if (['txt', 'md', 'csv', 'html'].includes(format)) {
    let encoding = 'utf-8';
    if (bytes[0] === 0xff && bytes[1] === 0xfe) encoding = 'utf-16le';
    else if (bytes[0] === 0xfe && bytes[1] === 0xff) encoding = 'utf-16be';
    let text;
    try { text = new TextDecoder(encoding, { fatal: true }).decode(bytes); }
    catch { text = new TextDecoder('gb18030', { fatal: true }).decode(bytes); }
    if (text.includes('\0')) throw new Error('文件包含二进制数据，无法作为文本读取。');
    if (text.length > 5000000) throw new Error('文本超过 500 万字符，请拆分后重发。');
    return text.replace(/^\uFEFF/, '').replaceAll('\r\n', '\n');
  }
  const parsed = new Promise((resolve, reject) => {
    const child = fork(fileURLToPath(new URL('./parser-worker.js', import.meta.url)), [], {
      stdio: ['ignore', 'ignore', 'ignore', 'ipc'], execArgv: ['--max-old-space-size=512'],
      env: { PATH: process.env.PATH || '', LANG: 'en_US.UTF-8' },
    });
    let settled = false;
    const done = (error, text) => {
      if (settled) return;
      settled = true; clearTimeout(timer); child.kill('SIGKILL');
      error ? reject(error) : resolve(text);
    };
    const timer = setTimeout(() => done(new Error('解析超过 45 秒，请拆分文件或另存后重试。')), timeoutMs);
    child.on('error', () => done(new Error('本地解析器启动失败，请检查插件依赖是否完整安装。')));
    child.on('exit', () => done(new Error('本地解析器意外退出，请另存文件后重试。')));
    child.on('message', msg => {
      if (!msg.ok) return done(new Error(hints[msg.code] || '本地解析失败，请另存后重发。'));
      if (!msg.text?.trim()) return done(new Error('未发现可读取的文字；扫描图片、图表和嵌入对象可能需要另外处理。'));
      if (msg.text.length > 5000000) return done(new Error('提取内容超过 500 万字符，请拆分文件后重发。'));
      done(null, msg.text);
    });
    child.send({ data: bytes.toString('base64'), format }, error => { if (error) done(new Error('本地解析器通信失败。')); });
  });
  try { return await parsed; }
  catch (error) {
    if (format !== 'doc' || process.platform !== 'darwin' || bytes.subarray(0, 8).toString('hex') !== 'd0cf11e0a1b11ae1') throw error;
    // macOS's local converter handles OLE Word variants rejected by anydoc.
    return new Promise((resolve, reject) => {
      const child = spawn('/usr/bin/textutil', ['-convert', 'txt', '-stdout', '-stdin', '-format', 'doc'], { stdio: ['pipe', 'pipe', 'ignore'], env: { PATH: '/usr/bin:/bin', LANG: 'en_US.UTF-8' } });
      let output = ''; let ended = false;
      const timer = setTimeout(() => { ended = true; child.kill('SIGKILL'); reject(new Error('旧版 Word 解析超时，请另存为 DOCX 后重试。')); }, timeoutMs);
      child.stdout.setEncoding('utf8');
      child.stdout.on('data', chunk => { output += chunk; if (output.length > 5000000) { ended = true; clearTimeout(timer); child.kill('SIGKILL'); reject(new Error('旧版 Word 内容过长，请拆分后重试。')); } });
      child.on('error', () => { clearTimeout(timer); if (!ended) { ended = true; reject(error); } });
      child.on('exit', code => { clearTimeout(timer); if (ended) return; ended = true; code === 0 && output.trim() ? resolve(output) : reject(error); });
      child.stdin.on('error', () => {}); child.stdin.end(bytes);
    });
  }
}
