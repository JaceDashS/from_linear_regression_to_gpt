import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop\\rebuilt_pair.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));

// Remove the two old direct backprop slides, highest index first.
presentation.slides.getItem(26).delete(); // old slide 27: Backprop sends error backward
presentation.slides.getItem(23).delete(); // old slide 24: Backpropagation reuses the chain rule

const C = { navy: "#0F172A", muted: "#475569", blue: "#2563EB", pale: "#EFF6FF", line: "#CBD5E1", white: "#FFFFFF", red: "#DC2626", green: "#059669", light: "#F8FAFC" };
function clear(slide) {
  return slide.export({ format: "layout" }).then(async blob => {
    const layout = JSON.parse(await blob.text());
    for (const e of layout.elements) if (e.aid) presentation.resolve(e.aid).delete();
    slide.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 1280, height: 720 }, fill: C.white, line: { style: "solid", fill: "none", width: 0 } });
  });
}
function text(slide, value, position, style = {}) {
  const shape = slide.shapes.add({ geometry: "rect", position, fill: { color: "FFFFFF", transparency: 100 }, line: { style: "solid", fill: "none", width: 0 } });
  shape.text = value;
  shape.text.style = { typeface: "Calibri", fontSize: 18, color: C.navy, ...style };
  if (style.verticalAlignment) shape.text.verticalAlignment = style.verticalAlignment;
  return shape;
}
function box(slide, value, position, opts = {}) {
  const shape = slide.shapes.add({ geometry: opts.geometry ?? "roundRect", position, fill: opts.fill ?? C.white, line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.3 } });
  if (value) { shape.text = value; shape.text.style = { typeface: "Calibri", fontSize: opts.fontSize ?? 18, color: opts.color ?? C.navy, bold: opts.bold ?? false, alignment: opts.alignment ?? "center" }; shape.text.verticalAlignment = "middle"; }
  return shape;
}
function arrow(slide, value, position, opts = {}) { return box(slide, value, position, { geometry: "rightArrow", fill: opts.fill ?? C.pale, line: opts.line ?? C.blue, color: opts.color ?? C.blue, fontSize: opts.fontSize ?? 18, bold: true }); }
function title(slide, value) { text(slide, value, { left: 66, top: 46, width: 1040, height: 46 }, { fontSize: 29, bold: true, color: C.navy }); }
function footer(slide, page) { text(slide, String(page), { left: 1166, top: 660, width: 48, height: 22 }, { fontSize: 13, color: "#64748B", alignment: "right" }); }
function note(slide, value) { slide.notes = `[Sources]\n${value}`; }

// Slide 22: simple chain rule.
const chain = presentation.slides.getItem(21);
await clear(chain);
title(chain, "The chain rule: multiply local changes");
text(chain, "If a parameter changes w → z → loss, the total effect is the product of each local effect.", { left: 68, top: 102, width: 1110, height: 30 }, { fontSize: 18, color: C.muted });
box(chain, "w\nweight", { left: 122, top: 238, width: 190, height: 105 }, { fill: "#F8FAFC", line: C.line, fontSize: 22, bold: true });
arrow(chain, "z = wx + b", { left: 370, top: 251, width: 245, height: 78 }, { fontSize: 19 });
arrow(chain, "L = loss(z)", { left: 690, top: 251, width: 245, height: 78 }, { fontSize: 19, fill: "#F0FDF4", line: C.green, color: C.green });
text(chain, "forward dependency", { left: 383, top: 346, width: 510, height: 25 }, { fontSize: 16, color: C.muted, alignment: "center" });
box(chain, "dL/dw  =  dL/dz  ×  dz/dw", { left: 210, top: 435, width: 860, height: 82 }, { fill: C.pale, line: C.blue, lineWidth: 1.8, fontSize: 29, bold: true, color: C.blue });
text(chain, "total change       =       downstream sensitivity       ×       local change", { left: 240, top: 535, width: 800, height: 25 }, { fontSize: 16, color: C.muted, alignment: "center" });
footer(chain, 22); note(chain, "Educational synthesis based on the standard chain rule in calculus.");

// Slide 23: backprop is the same calculation run backward through a neuron.
const backprop = presentation.slides.getItem(22);
await clear(backprop);
title(backprop, "Backpropagation applies the chain rule backward");
text(backprop, "Forward pass computes values; backward pass carries the loss gradient to every parameter.", { left: 68, top: 102, width: 1110, height: 30 }, { fontSize: 18, color: C.muted });
const xs = [74, 290, 506, 722, 938];
const labels = ["w", "z = wx + b", "h = σ(z)", "ŷ", "L(ŷ, y)"];
const fills = ["#F8FAFC", C.pale, "#F5F3FF", "#FFF7ED", "#FEF2F2"];
const lines = [C.line, C.blue, "#7C3AED", "#EA580C", C.red];
for (let i = 0; i < labels.length; i++) box(backprop, labels[i], { left: xs[i], top: 225, width: 168, height: 76 }, { fill: fills[i], line: lines[i], fontSize: i === 1 || i === 2 ? 19 : 22, bold: true });
for (let i = 0; i < 4; i++) arrow(backprop, "", { left: xs[i] + 170, top: 245, width: 44, height: 36 }, { fill: "#FFFFFF", line: C.line, color: C.line, fontSize: 1 });
text(backprop, "forward: compute prediction and loss", { left: 285, top: 318, width: 685, height: 26 }, { fontSize: 16, color: C.muted, alignment: "center" });
for (let i = 3; i >= 0; i--) { const a = box(backprop, "", { left: xs[i] + 170, top: 376, width: 44, height: 36 }, { geometry: "leftArrow", fill: "#FEE2E2", line: C.red }); }
text(backprop, "backward: send ∂L/∂(·) from loss to weight", { left: 285, top: 426, width: 685, height: 26 }, { fontSize: 16, color: C.red, alignment: "center", bold: true });
box(backprop, "∂L/∂w = (∂L/∂ŷ) × (∂ŷ/∂h) × (∂h/∂z) × (∂z/∂w)", { left: 117, top: 488, width: 1046, height: 67 }, { fill: "#FFF7ED", line: "#FB923C", lineWidth: 1.6, fontSize: 23, bold: true, color: "#9A3412" });
text(backprop, "Use this gradient to update the weight:   w ← w − η · ∂L/∂w", { left: 210, top: 580, width: 860, height: 28 }, { fontSize: 18, color: C.navy, alignment: "center", bold: true });
footer(backprop, 23); note(backprop, "Educational synthesis based on standard backpropagation and gradient descent.");

// Re-number footer text after deletion.
for (let i = 0; i < presentation.slides.items.length; i++) {
  const slide = presentation.slides.getItem(i);
  const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [x, y, w, h] = e.bbox ?? [];
    if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) {
      const f = presentation.resolve(e.aid); f.text = String(i + 1); f.text.style = { typeface: "Calibri", fontSize: 13, color: "#64748B", alignment: "right" };
    }
  }
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
for (const slideNum of [21,22,23,24]) { const slide = presentation.slides.getItem(slideNum - 1); const png = await slide.export({ format: "png", scale: 1 }); await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop\\slide-${slideNum}.png`, new Uint8Array(await png.arrayBuffer())); }
console.log(`slides=${presentation.slides.items.length}`);
