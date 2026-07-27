import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\cnn_rnn_reworked.pptx";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB", green: "#059669", orange: "#D97706", red: "#DC2626", paleBlue: "#EAF2FF", paleGreen: "#EAFBF3", paleOrange: "#FFF4E5", line: "#CBD5E1" };

function addText(slide, value, position, style = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  s.text = value;
  s.text.style = { typeface: "Calibri", fontSize: style.fontSize ?? 20, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  s.text.verticalAlignment = style.verticalAlignment ?? "middle";
  return s;
}
function addBox(slide, position, fill, line = C.line) {
  return slide.shapes.add({ geometry: "roundRect", position, fill, line: { style: "solid", fill: line, width: 1 }, borderRadius: "rounded-xl" });
}
async function resetDuplicate(slide, kicker, title) {
  let layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [, y] = e.bbox ?? [];
    if (e.aid && y >= 110 && y <= 642) p.resolve(e.aid).delete();
  }
  for (const image of [...slide.images.items]) image.delete();
  layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [x, y] = e.bbox ?? [];
    if (!e.aid || x < 60 || x > 90) continue;
    const shape = p.resolve(e.aid);
    if (y >= 30 && y <= 60) { shape.text = kicker; shape.text.style = { typeface: "Calibri", fontSize: 16, bold: true, color: C.blue, alignment: "left" }; }
    if (y >= 60 && y <= 112) { shape.text = title; shape.text.style = { typeface: "Calibri", fontSize: 36, bold: true, color: C.ink, alignment: "left" }; shape.text.verticalAlignment = "middle"; }
  }
}

// Insert an explanation directly after the existing CNN overview.
const cnnDetail = p.slides.getItem(22).duplicate();
await resetDuplicate(cnnDetail, "Vision", "CNNs detect local patterns in images");
addText(cnnDetail, "Use: vision tasks such as image classification, object detection, and segmentation.", { left: 74, top: 132, width: 1120, height: 30 }, { fontSize: 19, color: C.muted });
addBox(cnnDetail, { left: 78, top: 202, width: 362, height: 326 }, "#F8FAFC");
addText(cnnDetail, "A small filter slides across the image", { left: 104, top: 224, width: 310, height: 34 }, { fontSize: 21, bold: true });
// Image patch grid
for (let r = 0; r < 3; r++) for (let c = 0; c < 3; c++) {
  const color = (r === 1 && c === 1) || (r === 0 && c === 2) ? "#2563EB" : "#DCE8FF";
  cnnDetail.shapes.add({ geometry: "rect", position: { left: 125 + c * 44, top: 287 + r * 44, width: 38, height: 38 }, fill: color, line: { style: "solid", fill: "#FFFFFF", width: 1 } });
}
addText(cnnDetail, "image patch", { left: 130, top: 428, width: 112, height: 25 }, { fontSize: 16, color: C.muted, alignment: "center" });
addText(cnnDetail, "×", { left: 260, top: 317, width: 32, height: 34 }, { fontSize: 27, bold: true, color: C.muted, alignment: "center" });
for (let r = 0; r < 2; r++) for (let c = 0; c < 2; c++) cnnDetail.shapes.add({ geometry: "rect", position: { left: 307 + c * 38, top: 300 + r * 38, width: 32, height: 32 }, fill: (r + c) % 2 ? "#93C5FD" : "#1D4ED8", line: { style: "solid", fill: "#FFFFFF", width: 1 } });
addText(cnnDetail, "filter", { left: 300, top: 428, width: 90, height: 25 }, { fontSize: 16, color: C.muted, alignment: "center" });
addBox(cnnDetail, { left: 491, top: 202, width: 655, height: 326 }, "#F8FAFC");
addText(cnnDetail, "How it works", { left: 525, top: 224, width: 200, height: 32 }, { fontSize: 21, bold: true });
addText(cnnDetail, "1", { left: 529, top: 287, width: 32, height: 32 }, { fontSize: 20, bold: true, color: C.blue, alignment: "center" });
addText(cnnDetail, "Convolution: the same filter scans every local region.", { left: 576, top: 282, width: 520, height: 40 }, { fontSize: 19 });
addText(cnnDetail, "2", { left: 529, top: 350, width: 32, height: 32 }, { fontSize: 20, bold: true, color: C.blue, alignment: "center" });
addText(cnnDetail, "Activation map: strong matches mark useful features.", { left: 576, top: 345, width: 520, height: 40 }, { fontSize: 19 });
addText(cnnDetail, "3", { left: 529, top: 413, width: 32, height: 32 }, { fontSize: 20, bold: true, color: C.blue, alignment: "center" });
addText(cnnDetail, "Deeper layers combine edges → textures → objects.", { left: 576, top: 408, width: 520, height: 40 }, { fontSize: 19 });
addText(cnnDetail, "The key idea: reuse one set of weights wherever the pattern may appear.", { left: 118, top: 570, width: 1040, height: 28 }, { fontSize: 19, bold: true, color: C.blue, alignment: "center" });
cnnDetail.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard convolutional neural-network concepts."); cnnDetail.speakerNotes.setVisible(true);

// Insert a comparison immediately after the existing RNN/LSTM/GRU visual.
const recurrentDetail = p.slides.getItem(24).duplicate();
await resetDuplicate(recurrentDetail, "Sequences", "RNN, LSTM, and GRU: what each adds");
addText(recurrentDetail, "All pass information through time; gates improve what can be remembered.", { left: 74, top: 132, width: 1120, height: 30 }, { fontSize: 19, color: C.muted });
const cards = [
  { x: 76, name: "RNN", fill: C.paleBlue, accent: C.blue, role: "Carries one hidden state\nfrom token to token.", limit: "Limit: long-range context\ncan fade or vanish." },
  { x: 440, name: "LSTM", fill: C.paleGreen, accent: C.green, role: "Adds a memory cell plus\ninput / forget / output gates.", limit: "Trade-off: stronger memory,\nmore parameters and compute." },
  { x: 804, name: "GRU", fill: C.paleOrange, accent: C.orange, role: "Uses update and reset gates\nwith a simpler state.", limit: "Trade-off: lighter than LSTM,\nbut still sequential." },
];
for (const c of cards) {
  addBox(recurrentDetail, { left: c.x, top: 208, width: 328, height: 332 }, c.fill, c.accent);
  addText(recurrentDetail, c.name, { left: c.x + 28, top: 234, width: 270, height: 38 }, { fontSize: 28, bold: true, color: c.accent });
  addText(recurrentDetail, c.role, { left: c.x + 28, top: 302, width: 272, height: 68 }, { fontSize: 19 });
  recurrentDetail.shapes.add({ geometry: "line", position: { left: c.x + 28, top: 396, width: 270, height: 0 }, line: { style: "solid", fill: c.accent, width: 1 } });
  addText(recurrentDetail, c.limit, { left: c.x + 28, top: 420, width: 272, height: 65 }, { fontSize: 18, bold: true, color: C.red });
}
addText(recurrentDetail, "Shared limitation: each step depends on the previous step, so training and inference are hard to parallelize.", { left: 92, top: 576, width: 1092, height: 30 }, { fontSize: 18, bold: true, color: C.muted, alignment: "center" });
recurrentDetail.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard RNN, LSTM, and GRU architectures."); recurrentDetail.speakerNotes.setVisible(true);

// The legacy slides repeat the new RNN sequence/forgetting explanation.
for (let i = p.slides.items.length - 1; i >= 0; i--) {
  const s = p.slides.getItem(i); const layout = JSON.parse(await (await s.export({ format: "layout" })).text());
  const allText = layout.elements.filter(e => e.text).map(e => e.text).join(" ");
  if (allText.includes("RNNs read one token at a time") || allText.includes("Long memory fades")) s.delete();
}

for (let i = 0; i < p.slides.items.length; i++) {
  const s = p.slides.getItem(i); const elems = JSON.parse(await (await s.export({ format: "layout" })).text()).elements;
  for (const e of elems) { const [x, y, w, h] = e.bbox ?? []; if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) { const f = p.resolve(e.aid); f.text = String(i + 1); f.text.style = { typeface: "Calibri", fontSize: 13, color: C.muted, alignment: "right" }; } }
}
for (const n of [23, 24, 25, 26, 29, 30]) { const s = p.slides.getItem(n - 1); const b = await p.export({ slide: s, format: "png", scale: 1 }); await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\slide-${n}.png`, new Uint8Array(await b.arrayBuffer())); }
const pptx = await PresentationFile.exportPptx(p); await pptx.save(output); console.log(`slides=${p.slides.items.length}`);
