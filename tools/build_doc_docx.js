// Editable Word version of a Markdown document (the leadership explainer).
// Usage: NODE_PATH=<node_modules> node build_doc_docx.js in.md out.docx
// Handles the Markdown subset the explainer uses: # / ## / ### headings, paragraphs
// with **bold**, *italic*, `code` and [links](url), pipe tables, - bullets,
// 1. numbered lists, ``` code blocks, ![images](path) and --- rules.
const fs = require("fs"), path = require("path");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType,
        ExternalHyperlink, ImageRun, LevelFormat, BorderStyle, ShadingType, Footer, PageNumber } = require("docx");
const sizeOf = (() => { const m = require("image-size"); return m.imageSize || m.default || m; })();

const src = path.resolve(process.argv[2]), out = path.resolve(process.argv[3]);
const base = path.dirname(src);
const lines = fs.readFileSync(src, "utf8").split("\n");

const BODY = "Calibri", HEAD = "Cambria", INK = "15202B", MUTED = "4A5261", ACCENT = "A8560A", LINK = "1F6B78", RULE = "C9C4B8", SHADE = "ECE8DD";
const PAGE_W = 11906, MARGIN = 1134, CONTENT = PAGE_W - 2 * MARGIN;   // A4, DXA

// ---- inline markdown -> runs
function inline(text, extra = {}) {
  const runs = [];
  const re = /(\*\*.+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\)|https?:\/\/[^\s)|`]+|\*[^*]+\*)/g;
  let last = 0, m;
  const push = (t, o) => { if (t) runs.push(new TextRun(Object.assign({ text: t, font: BODY }, extra, o || {}))); };
  while ((m = re.exec(text))) {
    push(text.slice(last, m.index));
    const tok = m[0];
    if (tok.startsWith("http")) {                                  // bare URL -> clickable link
      const url = tok.replace(/[.,;:]+$/, ""); const trail = tok.slice(url.length);
      runs.push(new ExternalHyperlink({ link: url, children: [new TextRun(Object.assign({ text: url, font: BODY, color: LINK, underline: {} }, extra))] }));
      push(trail);
    }
    else if (tok.startsWith("**")) push(tok.slice(2, -2), { bold: true });
    else if (tok.startsWith("`")) push(tok.slice(1, -1), { font: "Courier New", size: (extra.size || 22) - 2, shading: { type: ShadingType.CLEAR, fill: "F2F0EA", color: "auto" } });
    else if (tok.startsWith("[")) {
      const mm = tok.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
      runs.push(new ExternalHyperlink({ link: mm[2], children: [new TextRun(Object.assign({ text: mm[1], font: BODY, color: LINK, underline: {} }, extra))] }));
    } else push(tok.slice(1, -1), { italics: true });
    last = m.index + tok.length;
  }
  push(text.slice(last));
  return runs;
}
const para = (text, o = {}) => new Paragraph(Object.assign({ children: inline(text, o.run || {}), spacing: { after: 120, line: 276 } }, o.p || {}));

// ---- blocks
const children = [];
let i = 0, numInstance = 0, sawTitle = false;
function cellWidths(rows) {
  const n = rows[0].length;
  const w = Array.from({ length: n }, (_, c) => Math.min(60, Math.max(6, ...rows.map(r => (r[c] || "").length))));
  const sum = w.reduce((a, b) => a + b, 0);
  return w.map(x => Math.max(Math.round(CONTENT * 0.12), Math.round(CONTENT * x / sum)));
}
function tableBlock(rowsText) {
  const rows = rowsText.map(l => l.replace(/^\||\|$/g, "").split("|").map(c => c.trim()));
  const header = rows[0], body = rows.slice(2);                       // rows[1] is the --- separator
  const widths = cellWidths([header, ...body]);
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (cells, isHead) => new TableRow({ tableHeader: isHead, children: cells.map((c, k) => new TableCell({
    width: { size: widths[k], type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 90, right: 90 },
    shading: isHead ? { type: ShadingType.CLEAR, fill: SHADE, color: "auto" } : undefined,
    children: [new Paragraph({ children: inline(c, { size: 19, bold: isHead || undefined }), spacing: { after: 0, line: 252 } })] })) });
  children.push(new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: [mk(header, true), ...body.map(r => mk(r, false))] }));
  children.push(new Paragraph({ spacing: { after: 120 } }));
}
while (i < lines.length) {
  const line = lines[i];
  if (!line.trim()) { i++; continue; }
  if (line.startsWith("# ") && !sawTitle) { sawTitle = true; children.push(new Paragraph({ heading: HeadingLevel.TITLE, children: inline(line.slice(2), { font: HEAD }) })); i++; continue; }
  if (line.startsWith("## ")) { children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: inline(line.slice(3), { font: HEAD }) })); i++; continue; }
  if (line.startsWith("### ")) { children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: inline(line.slice(4), { font: HEAD }) })); i++; continue; }
  if (/^-{3,}$/.test(line.trim())) { children.push(new Paragraph({ spacing: { before: 120, after: 240 }, border: { bottom: { color: RULE, space: 1, style: BorderStyle.SINGLE, size: 6 } } })); i++; continue; }
  if (line.startsWith("```")) {
    const code = []; i++;
    while (i < lines.length && !lines[i].startsWith("```")) code.push(lines[i++]);
    i++;
    code.forEach((c, k) => children.push(new Paragraph({ children: [new TextRun({ text: c || " ", font: "Courier New", size: 18 })],
      shading: { type: ShadingType.CLEAR, fill: "F2F0EA", color: "auto" }, spacing: { after: k === code.length - 1 ? 160 : 0, line: 240 } })));
    continue;
  }
  if (line.startsWith("|")) { const rows = []; while (i < lines.length && lines[i].startsWith("|")) rows.push(lines[i++]); tableBlock(rows); continue; }
  if (/^!\[/.test(line)) {
    const m = line.match(/^!\[([^\]]*)\]\(([^)]+)\)/);
    const p = path.resolve(base, m[2]); const buf = fs.readFileSync(p); const d = sizeOf(buf);
    const w = 600, h = Math.round(d.height * w / d.width);
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, children: [new ImageRun({ type: path.extname(p).slice(1) === "jpg" ? "jpg" : "png", data: buf, transformation: { width: w, height: h }, altText: { title: m[1], description: m[1], name: path.basename(p) } })] }));
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: m[1], font: BODY, size: 18, italics: true, color: MUTED })] }));
    i++; continue;
  }
  if (/^- /.test(line)) {
    while (i < lines.length && /^- /.test(lines[i])) {
      children.push(new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: inline(lines[i].slice(2)), spacing: { after: 60, line: 264 } })); i++;
    }
    children.push(new Paragraph({ spacing: { after: 60 } })); continue;
  }
  if (/^\d+\. /.test(line)) {
    numInstance++;
    while (i < lines.length && /^\d+\. /.test(lines[i])) {
      children.push(new Paragraph({ numbering: { reference: "numbers", level: 0, instance: numInstance }, children: inline(lines[i].replace(/^\d+\. /, "")), spacing: { after: 60, line: 264 } })); i++;
    }
    children.push(new Paragraph({ spacing: { after: 60 } })); continue;
  }
  // paragraph: join consecutive plain lines
  const buf = [line];
  i++;
  while (i < lines.length && lines[i].trim() && !/^(#|\||- |\d+\. |```|!\[|-{3,}$)/.test(lines[i])) buf.push(lines[i++]);
  children.push(para(buf.join(" ")));
}

const doc = new Document({
  creator: "Shreenath (Shree)", title: "Ami: hardening a retail customer-support AI agent for production", description: "L4 capstone project explainer",
  styles: {
    default: { document: { run: { font: BODY, size: 22, color: INK } } },
    paragraphStyles: [
      { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: HEAD, size: 44, bold: true, color: INK }, paragraph: { spacing: { before: 0, after: 200 } } },
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: HEAD, size: 32, bold: true, color: ACCENT }, paragraph: { spacing: { before: 360, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { font: HEAD, size: 26, bold: true, color: INK }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
      { id: "Hyperlink", name: "Hyperlink", basedOn: "Normal", run: { color: LINK, underline: {} } }] },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: "Ami · L4 capstone project · October 2026 · page ", font: BODY, size: 16, color: MUTED }),
      new TextRun({ children: [PageNumber.CURRENT], font: BODY, size: 16, color: MUTED })] })] }) },
    children }] });

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(out, buf); console.log("wrote", out, buf.length, "bytes,", children.length, "blocks"); });
