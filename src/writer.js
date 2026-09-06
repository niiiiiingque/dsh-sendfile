import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx';
import ExcelJS from 'exceljs';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { extOf, OUTPUT_EXTS } from './protocol.js';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml = text => new DOMParser({ onError(level, message) { if (level !== 'warning') throw new Error(message); } }).parseFromString(text, 'application/xml');
const serialize = dom => new XMLSerializer().serializeToString(dom);
const string = (value, limit = 200000) => { if (typeof value !== 'string' || value.length > limit) throw new Error('文本内容缺失或超出长度限制。'); return value; };
const primitive = value => { if (value !== null && !['string', 'number', 'boolean'].includes(typeof value)) throw new Error('表格单元格只接受文本、数字、布尔值或 null。'); if (typeof value === 'string' && value.length > 32767) throw new Error('单元格文本过长。'); return value; };

export async function createDocument(name, spec) {
  const format = extOf(name);
  if (!OUTPUT_EXTS.includes(format)) throw new Error('可创建 docx、xlsx、pptx、md、txt、csv、html；PDF 可通过打印导出或现有工具生成后登记。');
  if (['txt', 'md', 'csv', 'html'].includes(format)) return Buffer.from(string(spec.text), 'utf8');
  if (format === 'docx') {
    const body = string(spec.text);
    const paragraphs = body.split('\n').map(text => new Paragraph({ children: [new TextRun({ text, font: '宋体', size: 24 })], spacing: { after: 120, line: 360 } }));
    if (spec.title) paragraphs.unshift(new Paragraph({ text: string(spec.title, 200), heading: HeadingLevel.TITLE, spacing: { after: 240 } }));
    const doc = new Document({ creator: 'DSH 发文件', styles: { default: { document: { run: { font: '宋体', size: 24 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } }, children: paragraphs }] });
    return Packer.toBuffer(doc);
  }
  if (format === 'xlsx') {
    if (!Array.isArray(spec.sheets) || !spec.sheets.length || spec.sheets.length > 30) throw new Error('请提供 1 至 30 个 sheets，每个包含 name 和 rows。');
    const book = new ExcelJS.Workbook(); let cells = 0;
    for (const data of spec.sheets) {
      if (!Array.isArray(data.rows) || data.rows.length > 20000) throw new Error('每张表应包含 rows 数组，最多 20,000 行。');
      const sheet = book.addWorksheet(string(data.name, 31));
      for (const row of data.rows) {
        if (!Array.isArray(row) || row.length > 200 || (cells += row.length) > 200000) throw new Error('表格超出 200 列或总计 200,000 单元格限制。');
        sheet.addRow(row.map(primitive));
      }
      sheet.views = [{ state: 'frozen', ySplit: 1 }];
      sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34675C' } };
      for (const column of sheet.columns) column.width = 20;
      sheet.eachRow(row => { row.alignment = { vertical: 'top', wrapText: true }; });
    }
    book.calcProperties.fullCalcOnLoad = true;
    return Buffer.from(await book.xlsx.writeBuffer());
  }
  if (!Array.isArray(spec.slides) || spec.slides.length < 1 || spec.slides.length > 50) throw new Error('请提供 1 至 50 页 slides，每页包含 title、text。');
  const deck = new PptxGenJS(); deck.layout = 'LAYOUT_WIDE'; deck.author = 'DSH 发文件';
  for (const data of spec.slides) {
    const title = string(data.title, 50); const body = string(data.text, 600);
    const slide = deck.addSlide(); slide.background = { color: 'F5F4EF' };
    slide.addText(title, { x: 0.6, y: 0.5, w: 12, h: 1, fontFace: 'PingFang SC', fontSize: 30, bold: true, color: '183C35', breakLine: false, fit: 'shrink' });
    slide.addText(body, { x: 0.65, y: 1.85, w: 12, h: 4.8, fontFace: 'PingFang SC', fontSize: 22, color: '263B36', valign: 'top', fit: 'shrink', paraSpaceAfterPt: 14 });
  }
  return Buffer.from(await deck.write({ outputType: 'nodebuffer' }));
}

function replaceInParagraph(paragraph, replacements, counts) {
  const segments = []; let text = '';
  function walk(node) {
    if (node.namespaceURI === W && node.localName === 't') { const value = node.textContent; segments.push({ node, start: text.length, end: text.length + value.length }); text += value; return; }
    if (node.namespaceURI === W && ['tab', 'br', 'cr'].includes(node.localName)) { text += node.localName === 'tab' ? '\t' : '\n'; return; }
    for (let child = node.firstChild; child; child = child.nextSibling) walk(child);
  }
  walk(paragraph);
  // Reject overlapping edits; every edit refers to the original text.
  const edits = [];
  replacements.forEach((replacement, i) => {
    let offset = 0; let start;
    while ((start = text.indexOf(replacement.find, offset)) >= 0) {
      edits.push({ start, end: start + replacement.find.length, replacement: replacement.replace }); counts[i]++; offset = start + replacement.find.length;
    }
  });
  edits.sort((a, b) => a.start - b.start);
  for (let i = 1; i < edits.length; i++) if (edits[i].start < edits[i - 1].end) throw new Error('替换范围互相重叠，请分次修改。');
  for (const edit of edits.reverse()) {
    const affected = segments.filter(segment => segment.end > edit.start && segment.start < edit.end);
    for (let i = affected.length - 1; i >= 0; i--) {
      const segment = affected[i]; const a = Math.max(0, edit.start - segment.start); const b = Math.min(segment.end - segment.start, edit.end - segment.start);
      segment.node.textContent = segment.node.textContent.slice(0, a) + (i === 0 ? edit.replacement : '') + segment.node.textContent.slice(b);
      segment.node.setAttribute('xml:space', 'preserve');
    }
  }
}

async function editDocx(bytes, spec) {
  const replacements = spec.replacements;
  if (!Array.isArray(replacements) || !replacements.length || replacements.length > 100) throw new Error('请提供 1 至 100 条 replacements。');
  for (const r of replacements) {
    string(r.find, 10000); string(r.replace, 10000);
    if (!r.find || /[\r\n\t]/.test(r.find + r.replace) || !Number.isInteger(r.expectedCount) || r.expectedCount < 1) throw new Error('每条替换需含非空 find、replace 和正整数 expectedCount；此模式不跨段落或换行。');
  }
  const zip = await JSZip.loadAsync(bytes);
  const entries = Object.values(zip.files);
  if (entries.length > 4000 || entries.reduce((n, f) => n + (f._data?.uncompressedSize || 0), 0) > 128 * 1048576) throw new Error('DOCX 解压体积过大。');
  const body = zip.file('word/document.xml');
  if (!body) throw new Error('无效的 DOCX，缺少正文。');
  const doc = xml(await body.async('string'));
  const counts = replacements.map(() => 0);
  for (const p of Array.from(doc.getElementsByTagNameNS(W, 'p'))) replaceInParagraph(p, replacements, counts);
  replacements.forEach((r, i) => { if (counts[i] !== r.expectedCount) throw new Error(`替换“${r.find}”预期 ${r.expectedCount} 处，实际 ${counts[i]} 处；未保存修改。`); });
  // Default document policy: all sections have no header or footer.
  for (const tag of ['headerReference', 'footerReference']) for (const node of Array.from(doc.getElementsByTagNameNS(W, tag))) node.parentNode.removeChild(node);
  zip.file('word/document.xml', serialize(doc));
  const removed = new Set();
  for (const key of Object.keys(zip.files)) if (/^word\/(?:header|footer)[^/]*\.xml$/.test(key) || /^word\/_rels\/(?:header|footer)[^/]*\.xml\.rels$/.test(key)) { removed.add('/' + key); zip.remove(key); }
  for (const key of ['word/_rels/document.xml.rels', '[Content_Types].xml']) {
    if (!zip.file(key)) continue;
    const relationships = xml(await zip.file(key).async('string'));
    for (const node of Array.from(relationships.documentElement.childNodes)) {
      if (node.nodeType !== 1) continue;
      if (removed.has(node.getAttribute('PartName')) || /\/(header|footer)$/.test(node.getAttribute('Type') || '')) node.parentNode.removeChild(node);
    }
    zip.file(key, serialize(relationships));
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

export async function editDocument(bytes, name, spec) {
  const format = extOf(name);
  if (format === 'docx') return editDocx(bytes, spec);
  if (format === 'xlsx') {
    if (!Array.isArray(spec.cells) || !spec.cells.length || spec.cells.length > 1000) throw new Error('请提供 1 至 1,000 项 cells 修改。');
    const packageZip = await JSZip.loadAsync(bytes);
    if (Object.values(packageZip.files).reduce((n, f) => n + (f._data?.uncompressedSize || 0), 0) > 128 * 1048576) throw new Error('工作簿解压体积过大。');
    const book = new ExcelJS.Workbook(); await book.xlsx.load(bytes);
    const seen = new Set();
    for (const change of spec.cells) {
      const key = `${change.sheet}!${change.cell}`;
      if (seen.has(key)) throw new Error('同一单元格不能重复修改。'); seen.add(key);
      const sheet = book.getWorksheet(change.sheet);
      if (!sheet || !/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(change.cell)) throw new Error('工作表或单元格地址无效。');
      const cell = sheet.getCell(change.cell);
      if (!Object.hasOwn(change, 'expected') || JSON.stringify(cell.value) !== JSON.stringify(change.expected)) throw new Error(`${key} 的原值与 expected 不一致；未保存修改。`);
      cell.value = primitive(change.value);
    }
    for (const sheet of book.worksheets) sheet.eachRow(row => row.eachCell(cell => {
      const value = cell.value;
      if (value && typeof value === 'object' && (value.formula || value.sharedFormula)) {
        const { result: cachedResult, ...calculation } = value; cell.value = calculation;
      }
    }));
    book.calcProperties.fullCalcOnLoad = true;
    return Buffer.from(await book.xlsx.writeBuffer());
  }
  if (['txt', 'md', 'csv', 'html'].includes(format)) return createDocument(name, spec);
  throw new Error('此版可修改 DOCX 的文字、XLSX 的单元格及文本文件。旧格式、PDF 和 PPTX 请根据内容另建新文件。');
}
