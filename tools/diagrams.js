// Diagrams for the project explainer, drawn once from a spec and rendered two ways:
//   toSVG(spec)                      -> an SVG string (for the Word/PDF/Markdown versions)
//   toPptx(pres, slide, spec, x, y)  -> native, editable shapes on a pptxgenjs slide
// Canvas is a 1600 x 680 logical grid. Run directly to export SVGs:
//   node diagrams.js <outdir>
const PALETTE = { ink: "15202B", muted: "4A5261", cream: "ECE8DD", creamLine: "C9C4B8", orange: "A8560A", orangeLight: "F6E0B5",
  teal: "1F6B78", tealLight: "D9ECEF", navy: "1E2C3A", navyText: "F6F4EE", green: "1E5E33", greenLight: "D7EBD9", grey: "F2F0EA", white: "FFFFFF" };
const P = PALETTE;

// ---------- node helpers -------------------------------------------------
const box = (id, x, y, w, h, title, sub, o = {}) => Object.assign({ id, kind: "box", x, y, w, h, title, sub: sub || [], fill: P.cream, stroke: P.creamLine, color: P.ink }, o);
const dark = (id, x, y, w, h, title, sub, o = {}) => box(id, x, y, w, h, title, sub, Object.assign({ fill: P.navy, stroke: P.navy, color: P.navyText, subColor: "C9D1D9" }, o));
const hot = (id, x, y, w, h, title, sub, o = {}) => box(id, x, y, w, h, title, sub, Object.assign({ fill: P.orange, stroke: P.orange, color: P.white, subColor: "FBE9D2" }, o));
const cool = (id, x, y, w, h, title, sub, o = {}) => box(id, x, y, w, h, title, sub, Object.assign({ fill: P.tealLight, stroke: P.teal, color: P.ink }, o));
const group = (id, x, y, w, h, title, o = {}) => Object.assign({ id, kind: "group", x, y, w, h, title, fill: P.grey, stroke: P.creamLine, color: P.muted }, o);
const note = (id, x, y, w, h, lines, o = {}) => Object.assign({ id, kind: "note", x, y, w, h, title: "", sub: lines, fill: P.white, stroke: P.creamLine, color: P.muted }, o);
const edge = (pts, o = {}) => Object.assign({ pts, arrow: true, color: P.muted, width: 2.2 }, o);

// ---------- 1. system architecture ----------------------------------------
const architecture = { id: "architecture", title: "How Ami is built", nodes: [
  note("agents", 40, 40, 250, 165, ["How many agents? One.", "Ami, the support agent:", "one LLM, two planner", "strategies, seven tools.", "The LLM judge used in", "testing is a grader, not", "an agent."]),
  cool("customer", 40, 230, 250, 130, "Customer", ["Browser chat UI", "Log in · Agent internals panel", "Thumbs up / down"]),
  cool("staff", 40, 440, 250, 150, "Support staff", ["/logs dashboard:", "cost, latency, alerts, events", "Ticket queue in Jira"]),
  box("web", 350, 170, 260, 300, "Web server", ["web.py, plain Python HTTP", "", "Login → signed JWT, scope =", "this customer only", "Rate limit · audit log", "Sessions · feedback"]),
  group("core", 670, 40, 480, 600, "AI agent core: Ami, the one agent"),
  hot("policy", 700, 85, 420, 70, "Policy layer", ["check_input · guarded_run · check_output"]),
  dark("planner", 700, 185, 230, 125, "Planner, ReAct", ["Thought → Action →", "Observation, ≤ 6 steps", "(plan-and-execute alt.)"]),
  box("memory", 955, 185, 165, 125, "Memory", ["conversation", "working", "long-term", "answer cache"]),
  box("tools", 700, 340, 420, 95, "Tools (7), guardrails inside", ["find / get orders · track · cancel · return", "search_knowledge · escalate"]),
  box("profile", 700, 460, 230, 70, "Profile", ["the system prompt: who Ami is"]),
  box("router", 955, 460, 165, 70, "Model router", ["easy → cheap tier"]),
  cool("observe", 700, 555, 420, 65, "Observability", ["every call, tool and turn: tokens, cost, latency → alerts"]),
  dark("llm", 1220, 55, 340, 110, "LLM", ["gpt-5.6-terra through an", "OpenAI-compatible API:", "chat completions + function calling"]),
  box("kb", 1220, 200, 340, 120, "Knowledge base (RAG)", ["15 policy documents → chunks →", "local embeddings (bge-micro-v2)", "→ ChromaDB vector index"]),
  box("store", 1220, 355, 340, 75, "Order store", ["orders and customers, scope-filtered"]),
  box("tickets", 1220, 465, 340, 155, "Hand-off to humans", ["ESC / CXL / RMA tickets", "→ escalation matrix: department,", "priority, SLA, next steps", "→ Jira · webhook · email"]),
], edges: [
  edge([[290, 295], [350, 295]]), edge([[610, 295], [670, 295]]),
  edge([[930, 247], [955, 247]], { arrow: false }),
  edge([[1120, 120], [1220, 120]], { twoWay: true }),
  edge([[1120, 370], [1220, 370]], { arrow: false }), edge([[1170, 370], [1170, 260], [1220, 260]]), edge([[1170, 370], [1170, 392], [1220, 392]]), edge([[1170, 370], [1170, 540], [1220, 540]]),
  edge([[1390, 620], [1390, 655], [165, 655], [165, 590]], { dashed: true, color: P.teal, label: "ticket reaches a person", labelAt: [700, 672] }),
  edge([[700, 587], [620, 587], [620, 515], [290, 515]], { dashed: true, color: P.teal }),
  edge([[815, 155], [815, 185]], { arrow: false }), edge([[815, 310], [815, 340]], { arrow: false }), edge([[1037, 310], [1037, 340]], { arrow: false }),
] };

// ---------- 2. one turn, step by step --------------------------------------
const turn = { id: "turn", title: "One customer turn, step by step", nodes: [
  cool("msg", 40, 40, 210, 90, "Customer message", ["typed in the chat"]),
  box("inpol", 290, 40, 210, 90, "1 · Input policy", ["scrub card numbers,", "flag override attempts"]),
  box("classify", 540, 40, 210, 90, "2 · Classify", ["PUBLIC (policy) or", "PRIVATE (account)"]),
  box("checks", 790, 40, 210, 90, "3 · Gate and cache", ["login required?", "cached answer?"]),
  box("route", 1040, 40, 210, 90, "4 · Model routing", ["easy → cheap tier,", "hard → strong"]),
  box("mem", 1290, 40, 270, 90, "5 · Memory and prefetch", ["long-term memory read,", "orders loaded on login"]),
  group("loop", 540, 225, 1020, 215, "6 · The ReAct loop, up to six steps"),
  dark("llm", 570, 290, 170, 90, "LLM call", ["profile, transcript,", "memory, tool menu"]),
  box("ask", 780, 290, 170, 90, "Tool needed?", ["yes → guarded run", "no → answer"]),
  hot("guard", 990, 290, 170, 90, "guarded_run", ["scope · confirmation", "· eligibility"]),
  box("run", 1200, 290, 170, 90, "Tool runs", ["orders · knowledge", "· tickets"]),
  box("obs", 1400, 290, 150, 90, "Observation", ["result appended"]),
  note("exits", 40, 250, 460, 170, ["Exits before the loop:", "• anonymous account question →", "  login gate: \"please log in\"", "• repeated public question →", "  cached reply, 0 model calls"]),
  box("outpol", 540, 530, 210, 90, "7 · Output policy", ["mask any identifier the", "tools did not produce"]),
  cool("log", 790, 530, 210, 90, "8 · Observability", ["log tokens, cost, latency;", "check alert thresholds"]),
  box("write", 1040, 530, 210, 90, "9 · Memory write", ["what was learned about", "this customer"]),
  cool("reply", 1290, 530, 270, 90, "Reply to the customer", ["with the trace for the", "internals panel"]),
], edges: [
  edge([[250, 85], [290, 85]]), edge([[500, 85], [540, 85]]), edge([[750, 85], [790, 85]]), edge([[1000, 85], [1040, 85]]), edge([[1250, 85], [1290, 85]]),
  edge([[1425, 130], [1425, 180], [620, 180], [620, 290]]),
  edge([[740, 335], [780, 335]]), edge([[950, 335], [990, 335]]), edge([[1160, 335], [1200, 335]]), edge([[1370, 335], [1400, 335]]),
  edge([[1475, 290], [1475, 258], [700, 258], [700, 290]], { label: "back to the model with what it learned", labelAt: [1090, 250] }),
  edge([[865, 380], [865, 485], [645, 485], [645, 530]], { label: "answer ready", labelAt: [905, 470] }),
  edge([[750, 575], [790, 575]]), edge([[1000, 575], [1040, 575]]), edge([[1250, 575], [1290, 575]]),
  edge([[895, 130], [895, 180], [270, 180], [270, 250]], { dashed: true, color: P.teal }),
  edge([[270, 420], [270, 650], [1425, 650], [1425, 620]], { dashed: true, color: P.teal, label: "straight to the reply", labelAt: [850, 668] }),
] };

// ---------- 3. RAG pipeline ------------------------------------------------
const rag = { id: "rag", title: "Retrieval-augmented generation: the agent answers from the written rules", nodes: [
  group("index", 40, 40, 1520, 200, "Indexing: once, and again whenever a document changes"),
  box("docs", 70, 85, 330, 130, "15 written documents", ["policies · rules · tone · regulations", "(returns, cancellations, shipping,", "refunds, gift cards, privacy …)"]),
  box("chunk", 450, 85, 290, 130, "Chunk at headings", ["about 48 chunks, each", "labelled with its document", "and category"]),
  dark("embed", 790, 85, 320, 130, "Embed locally", ["bge-micro-v2, 17 MB, int8,", "runs on the server, no API cost;", "384 numbers per chunk"]),
  cool("chroma", 1160, 85, 370, 130, "ChromaDB vector index", ["nearest-neighbour search;", "rebuilt automatically when", "the documents change"]),
  group("answer", 40, 300, 1520, 200, "Answering: every policy question, every time"),
  cool("q", 70, 345, 250, 130, "Customer question", ["\"What is your return", "policy?\""]),
  dark("qe", 360, 345, 250, 130, "Embed the question", ["same local model,", "same 384 numbers"]),
  box("near", 650, 345, 270, 130, "Nearest chunks", ["the passages whose", "numbers are closest", "(search_knowledge tool)"]),
  hot("read", 960, 345, 270, 130, "LLM reads evidence", ["prompt = question +", "retrieved passages"]),
  cool("ans", 1270, 345, 260, 130, "Answer", ["quotes the written rule,", "not the model's memory"]),
  note("down", 40, 545, 740, 105, ["If the knowledge store is unavailable, the tool returns nothing and the agent says", "\"I can't look that up right now\" — it never invents a policy."]),
  note("prefetch", 820, 545, 740, 105, ["Delivery problems (\"says delivered but not here\", \"tracking hasn't moved\") look", "the policy up before the first model call:", "16 → 9 model calls, median turn 4.9 s → 2.1 s."]),
], edges: [
  edge([[400, 150], [450, 150]]), edge([[740, 150], [790, 150]]), edge([[1110, 150], [1160, 150]]),
  edge([[320, 410], [360, 410]]), edge([[610, 410], [650, 410]]), edge([[920, 410], [960, 410]]), edge([[1230, 410], [1270, 410]]),
  edge([[1345, 215], [1345, 270], [785, 270], [785, 345]], { dashed: true, color: P.teal, label: "search", labelAt: [1070, 262] }),
] };

// ---------- 4. the AI technology stack --------------------------------------
const chip = (x, y, w, text, o = {}) => Object.assign({ kind: "chip", x, y, w, h: 50, title: text, fill: P.white, stroke: P.creamLine, color: P.ink }, o);
function layer(y, title, subtitle, fill, stroke, color, chips) {
  const nodes = [Object.assign({ kind: "layer", x: 40, y, w: 1340, h: 96, title, sub: [subtitle], fill, stroke, color })];
  let x = 330;
  chips.forEach(([t, w, o]) => { nodes.push(chip(x, y + 23, w, t, o)); x += w + 10; });
  return nodes;
}
const navyChip = { fill: "2E3F50", stroke: "2E3F50", color: P.navyText };
const stack = { id: "stack", title: "The AI technology stack, and how the layers connect", nodes: [].concat(
  layer(30, "Experience", "what customers and staff see", P.tealLight, P.teal, P.ink,
    [["Chat UI", 130], ["Login gate", 140], ["Agent internals panel", 200], ["Thumbs up / down", 190], ["/logs dashboard", 190]]),
  layer(162, "AI agent: Ami", "decides, acts, remembers", P.orangeLight, P.orange, P.ink,
    [["Profile (prompt)", 170], ["ReAct planner", 150], ["Plan-and-execute", 170], ["Memory ×3 + cache", 180], ["7 guarded tools", 160], ["Policy layer", 130]]),
  layer(294, "AI services", "the models it calls", P.navy, P.navy, P.navyText,
    [["LLM: gpt-5.6-terra via an OpenAI-compatible API", 470, navyChip], ["Local embeddings: bge-micro-v2", 290, navyChip], ["LLM judge (evals only)", 230, navyChip]]),
  layer(426, "Data and knowledge", "what the agent reads and writes", P.cream, P.creamLine, P.ink,
    [["Knowledge base, 15 docs", 230], ["ChromaDB index", 150], ["Order store, scope-filtered", 250], ["Escalation matrix", 165], ["Trace & audit logs", 170]]),
  layer(558, "Platform, operations", "how it runs and is kept honest", P.grey, P.creamLine, P.ink,
    [["Python 3.13, stdlib HTTP", 230], ["JWT auth", 100], ["Vercel / Docker", 150], ["Jira · webhook · email", 210], ["pytest 415 · evals 33 · golden 28", 300]])
), edges: [
  edge([[1455, 126], [1455, 162]], { twoWay: true, color: P.teal, label: "chat ↔ replies", labelAt: [1475, 150], labelAnchor: "start" }),
  edge([[1455, 258], [1455, 294]], { twoWay: true, color: P.orange, label: "prompt ↔ answer", labelAt: [1475, 282], labelAnchor: "start" }),
  edge([[1455, 390], [1455, 426]], { twoWay: true, color: P.muted, label: "reads & writes", labelAt: [1475, 414], labelAnchor: "start" }),
  edge([[1455, 522], [1455, 558]], { arrow: false, color: P.muted, label: "runs & tests", labelAt: [1475, 546], labelAnchor: "start" }),
] };

const DIAGRAMS = { architecture, turn, rag, stack };

// ---------- SVG renderer ---------------------------------------------------
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
function toSVG(spec, W = 1600, H = 680) {
  const FONT = "Calibri, Carlito, Arial, Helvetica, sans-serif";
  const colors = [...new Set(spec.edges.map(e => e.color))];
  let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="${FONT}"><rect width="${W}" height="${H}" fill="#FFFFFF"/><defs>`;
  colors.forEach(c => { s += `<marker id="a${c}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="#${c}"/></marker>`; });
  s += `</defs>`;
  const text = (x, y, t, size, fill, o = {}) => `<text x="${x}" y="${y}" font-size="${size}" fill="#${fill}"${o.bold ? ' font-weight="bold"' : ""}${o.italic ? ' font-style="italic"' : ""} text-anchor="${o.anchor || "start"}">${esc(t)}</text>`;
  for (const n of spec.nodes) {
    if (n.kind === "group" || n.kind === "layer") {
      s += `<rect x="${n.x}" y="${n.y}" width="${n.w}" height="${n.h}" rx="14" fill="#${n.fill}" stroke="#${n.stroke}" stroke-width="1.5"${n.kind === "group" ? ' stroke-dasharray="7 5"' : ""}/>`;
      if (n.kind === "group") s += text(n.x + 16, n.y + 28, n.title, 21, n.color, { bold: true });
      else { s += text(n.x + 20, n.y + 48, n.title, 24, n.color, { bold: true }); s += text(n.x + 20, n.y + 76, n.sub[0], 17, n.color, { italic: true }); }
      continue;
    }
    if (n.kind === "chip") {
      s += `<rect x="${n.x}" y="${n.y}" width="${n.w}" height="${n.h}" rx="27" fill="#${n.fill}" stroke="#${n.stroke}" stroke-width="1.2"/>`;
      s += text(n.x + n.w / 2, n.y + 33, n.title, 18, n.color, { anchor: "middle" });
      continue;
    }
    const isNote = n.kind === "note";
    s += `<rect x="${n.x}" y="${n.y}" width="${n.w}" height="${n.h}" rx="${isNote ? 8 : 12}" fill="#${n.fill}" stroke="#${n.stroke}" stroke-width="${isNote ? 1.2 : 1.6}"/>`;
    const lines = (n.title ? [n.title] : []).concat(n.sub);
    const lh = isNote ? 24 : 22, titleH = n.title ? 30 : 0;
    const total = titleH + n.sub.length * lh;
    let y = n.y + (n.h - total) / 2 + (n.title ? 24 : 18);
    if (n.title) { s += text(n.x + n.w / 2, y, n.title, 24, n.color, { bold: true, anchor: "middle" }); y += 28; }
    n.sub.forEach(l => { s += text(isNote ? n.x + 16 : n.x + n.w / 2, y, l, isNote ? 19 : 19, n.subColor || n.color, { anchor: isNote ? "start" : "middle", italic: isNote }); y += lh; });
  }
  for (const e of spec.edges) {
    const d = e.pts.map((p, i) => (i ? "L" : "M") + p[0] + "," + p[1]).join(" ");
    s += `<path d="${d}" fill="none" stroke="#${e.color}" stroke-width="${e.width}"${e.dashed ? ' stroke-dasharray="8 6"' : ""}${e.arrow ? ` marker-end="url(#a${e.color})"` : ""}${e.twoWay ? ` marker-start="url(#a${e.color})"` : ""}/>`;
    if (e.label) s += `<text x="${e.labelAt[0]}" y="${e.labelAt[1]}" font-size="16" fill="#${e.color}" text-anchor="${e.labelAnchor || "middle"}" font-style="italic">${esc(e.label)}</text>`;
  }
  return s + `</svg>`;
}

// ---------- pptxgenjs renderer (native shapes) ------------------------------
function toPptx(pres, slide, spec, x0 = 0.6, y0 = 1.75, width = 12.13, canvasW = 1600, canvasH = 680) {
  const k = width / canvasW, pt = (px) => Math.max(9, Math.round(px * k * 72 * 10) / 10);
  const X = (px) => x0 + px * k, Y = (px) => y0 + px * k, S = (px) => px * k;
  const T = (t, o) => slide.addText(t, Object.assign({ margin: 0, isTextBox: true }, o));
  for (const n of spec.nodes) {
    const isGroup = n.kind === "group" || n.kind === "layer", isChip = n.kind === "chip", isNote = n.kind === "note";
    slide.addShape(pres.ShapeType.roundRect, { x: X(n.x), y: Y(n.y), w: S(n.w), h: S(n.h), fill: { color: n.fill },
      line: { color: n.stroke, width: isGroup ? 1 : 0.75, dashType: n.kind === "group" ? "dash" : "solid" }, rectRadius: isChip ? 0.2 : 0.08, objectName: n.title || n.id || n.kind });
    if (n.kind === "group") { T(n.title, { x: X(n.x + 14), y: Y(n.y + 8), w: S(n.w - 28), h: S(28), fontSize: pt(21), bold: true, color: n.color, valign: "top" }); continue; }
    if (n.kind === "layer") {
      T([{ text: n.title, options: { bold: true, fontSize: pt(24), breakLine: true } }, { text: n.sub[0], options: { italic: true, fontSize: pt(17) } }], { x: X(n.x + 18), y: Y(n.y + 14), w: S(290), h: S(n.h - 28), color: n.color, valign: "middle" }); continue;
    }
    if (isChip) { T(n.title, { x: X(n.x + 6), y: Y(n.y), w: S(n.w - 12), h: S(n.h), fontSize: pt(18), color: n.color, align: "center", valign: "middle" }); continue; }
    const runs = [];
    if (n.title) runs.push({ text: n.title, options: { bold: true, fontSize: pt(24), color: n.color, breakLine: n.sub.length > 0 } });
    n.sub.forEach((l, i) => runs.push({ text: l || " ", options: { fontSize: pt(19), italic: isNote, color: n.subColor || n.color, breakLine: i < n.sub.length - 1 } }));
    T(runs, { x: X(n.x + 10), y: Y(n.y + 4), w: S(n.w - 20), h: S(n.h - 8), align: isNote ? "left" : "center", valign: "middle", lineSpacingMultiple: 1.0 });
  }
  for (const e of spec.edges) {
    for (let i = 1; i < e.pts.length; i++) {
      const [ax, ay] = e.pts[i - 1], [bx, by] = e.pts[i];
      const last = i === e.pts.length - 1, first = i === 1;
      const o = { x: X(Math.min(ax, bx)), y: Y(Math.min(ay, by)), w: S(Math.abs(bx - ax)), h: S(Math.abs(by - ay)),
        line: { color: e.color, width: Math.max(1, e.width * 0.6), dashType: e.dashed ? "dash" : "solid" }, objectName: "connector" };
      if (bx < ax) o.flipH = true; if (by < ay) o.flipV = true;
      if (last && e.arrow) o.line.endArrowType = "triangle";
      if (first && e.twoWay) o.line.beginArrowType = "triangle";
      slide.addShape(pres.ShapeType.line, o);
    }
    if (e.label) T(e.label, { x: X(e.labelAnchor === "start" ? e.labelAt[0] : e.labelAt[0] - 220), y: Y(e.labelAt[1] - 18), w: S(e.labelAnchor === "start" ? 130 : 440), h: S(22), fontSize: pt(16), italic: true, color: e.color, align: e.labelAnchor === "start" ? "left" : "center", valign: "middle" });
  }
}

module.exports = { DIAGRAMS, toSVG, toPptx, PALETTE };

if (require.main === module) {
  const fs = require("fs"), path = require("path");
  const out = process.argv[2] || ".";
  fs.mkdirSync(out, { recursive: true });
  for (const [id, spec] of Object.entries(DIAGRAMS)) { fs.writeFileSync(path.join(out, `diagram-${id}.svg`), toSVG(spec)); console.log("wrote", path.join(out, `diagram-${id}.svg`)); }
}
