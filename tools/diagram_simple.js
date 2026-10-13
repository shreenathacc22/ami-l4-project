// The one-picture "how Ami works" diagram for the leadership talk: same spec format and
// renderers as tools/diagrams.js (toSVG for the guide/PDF, toPptx for native slide shapes).
//   node diagram_simple.js <outdir>   -> writes diagram-simple.svg
const { toSVG, PALETTE: P } = require("./diagrams.js");

const box = (id, x, y, w, h, title, sub, o = {}) => Object.assign({ id, kind: "box", x, y, w, h, title, sub: sub || [], fill: P.cream, stroke: P.creamLine, color: P.ink }, o);
const dark = (id, x, y, w, h, title, sub) => box(id, x, y, w, h, title, sub, { fill: P.navy, stroke: P.navy, color: P.navyText, subColor: "C9D1D9" });
const hot = (id, x, y, w, h, title, sub) => box(id, x, y, w, h, title, sub, { fill: P.orange, stroke: P.orange, color: P.white, subColor: "FBE9D2" });
const cool = (id, x, y, w, h, title, sub) => box(id, x, y, w, h, title, sub, { fill: P.tealLight, stroke: P.teal, color: P.ink });
const note = (id, x, y, w, h, lines) => ({ id, kind: "note", x, y, w, h, title: "", sub: lines, fill: P.white, stroke: P.creamLine, color: P.muted });
const edge = (pts, o = {}) => Object.assign({ pts, arrow: true, color: P.muted, width: 2.2 }, o);

const simple = { id: "simple", title: "How Ami works, in one picture", nodes: [
  cool("msg", 40, 50, 250, 110, "Customer message", ["\"Where is my order?\"", "\"Cancel it\" · \"Returns?\""]),
  box("safety", 320, 50, 230, 110, "1 · Safety check", ["card numbers removed,", "manipulation flagged"]),
  box("classify", 580, 50, 260, 110, "2 · Which kind?", ["general policy, or", "this customer's account?"]),
  box("policy", 580, 215, 260, 90, "Policy question", ["look up the store's", "own written rules"]),
  box("account", 580, 320, 260, 90, "Account question", ["log in first, then", "only your own orders"]),
  dark("brain", 880, 50, 280, 360, "3 · The AI decides", ["reads the question,", "the rules and your orders,", "then picks the next step:", "answer, look something up,", "or take an action"]),
  hot("guard", 1190, 50, 250, 360, "4 · Guardrails", ["no change without", "your \"yes\"", "", "rules enforced in code,", "not just in the prompt", "", "only your own data"]),
  cool("reply", 1190, 440, 250, 100, "5 · Reply", ["quotes the policy or", "confirms the action"]),
  box("human", 880, 440, 280, 100, "Hand-off to a person", ["a ticket to the right team,", "with the facts attached"]),
  note("one", 40, 215, 250, 325, ["One AI agent, not a", "free-talking chatbot.", "", "Every lookup and action", "goes through rules we", "wrote, and a person is", "always one step away."]),
  cool("log", 320, 580, 1120, 80, "Everything is recorded", ["cost and time per reply, which lookups and actions ran, errors; alerts fire when quality, cost or speed drift"]),
], edges: [
  edge([[290, 105], [320, 105]]), edge([[550, 105], [580, 105]]),
  edge([[710, 160], [710, 215]]),
  edge([[580, 130], [555, 130], [555, 365], [580, 365]]),
  edge([[840, 260], [880, 260]]), edge([[840, 365], [880, 365]]),
  edge([[1160, 230], [1190, 230]]),
  edge([[1315, 410], [1315, 440]]),
  edge([[1250, 410], [1250, 425], [1020, 425], [1020, 440]]),
] };

module.exports = { simple };

if (require.main === module) {
  const fs = require("fs"), path = require("path");
  const out = process.argv[2] || "."; fs.mkdirSync(out, { recursive: true });
  fs.writeFileSync(path.join(out, "diagram-simple.svg"), toSVG(simple)); console.log("wrote", path.join(out, "diagram-simple.svg"));
}
