import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\edited.pptx";
const outDir = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const C = { ink: "#0B0F19", muted: "#667085", line: "#CBD5E1", blue: "#2563EB", paleBlue: "#EFF6FF", red: "#EF4444", paleRed: "#FEF2F2", green: "#059669", paleGreen: "#ECFDF5", orange: "#F97316", paleOrange: "#FFF7ED", white: "#FFFFFF", panel: "#F8FAFC" };

function text(slide, value, position, style = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  s.text = value;
  s.text.style = { typeface: "Calibri", fontSize: style.fontSize ?? 20, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  if (style.verticalAlignment) s.text.verticalAlignment = style.verticalAlignment;
  return s;
}
function box(slide, value, position, opts = {}) {
  const geometry = opts.geometry ?? "roundRect";
  const s = slide.shapes.add({ geometry, position, fill: opts.fill ?? C.white, line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 } });
  if (value) {
    s.text = value;
    s.text.style = { typeface: "Calibri", fontSize: opts.fontSize ?? 22, bold: opts.bold ?? true, color: opts.color ?? C.ink, alignment: opts.alignment ?? "center" };
    s.text.verticalAlignment = "middle";
  }
  return s;
}
function arrow(slide, position, direction = "right", color = C.blue) {
  return box(slide, "", position, { geometry: direction === "left" ? "leftArrow" : "rightArrow", fill: direction === "left" ? C.paleRed : C.paleBlue, line: color, lineWidth: 1.4 });
}
async function clearBody(slide) {
  const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [, y] = e.bbox ?? [];
    if (e.aid && y >= 142 && y <= 642) p.resolve(e.aid).delete();
  }
  for (const image of [...slide.images.items]) image.delete();
}
async function setHeader(slide, kicker, heading) {
  const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [x, y] = e.bbox ?? [];
    if (!e.aid) continue;
    const shape = p.resolve(e.aid);
    if (x >= 60 && x <= 80 && y >= 35 && y <= 55) shape.text = kicker;
    if (x >= 60 && x <= 80 && y >= 58 && y <= 80) shape.text = heading;
  }
}
function panel(slide) { box(slide, "", { left: 64, top: 142, width: 1152, height: 500 }, { fill: C.panel, line: "#E2E8F0", lineWidth: 1 }); }
function notes(slide, sourceText) { slide.speakerNotes.textFrame.setText(`[Sources]\n${sourceText}`); slide.speakerNotes.setVisible(true); }

// Reuse the existing chain-rule template slide immediately after Common activation functions.
const chain = p.slides.getItem(19);
const backprop = chain.duplicate();

await clearBody(chain);
await setHeader(chain, "Calculus", "The chain rule multiplies local changes");
panel(chain);
text(chain, "For composed functions, follow the dependency path and multiply each local derivative.", { left: 150, top: 180, width: 980, height: 34 }, { fontSize: 22, color: C.muted, alignment: "center" });

// Arrows first, then nodes, so connectors remain behind the labels.
arrow(chain, { left: 356, top: 288, width: 92, height: 44 });
arrow(chain, { left: 718, top: 288, width: 92, height: 44 });
box(chain, "x", { left: 170, top: 265, width: 170, height: 90 }, { fontSize: 28, line: C.line });
box(chain, "u = g(x)", { left: 465, top: 265, width: 235, height: 90 }, { fontSize: 27, fill: C.paleBlue, line: C.blue, color: C.blue });
box(chain, "y = f(u)", { left: 828, top: 265, width: 235, height: 90 }, { fontSize: 27, fill: C.paleGreen, line: C.green, color: C.green });
box(chain, "dy/dx  =  dy/du  ×  du/dx", { left: 220, top: 420, width: 840, height: 82 }, { fontSize: 31, fill: C.white, line: C.blue, lineWidth: 1.8, color: C.ink });
text(chain, "total effect  =  downstream change  ×  local change", { left: 270, top: 520, width: 740, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });
text(chain, "Example: u = 3x + 1, y = u²  →  dy/dx = 2u × 3", { left: 260, top: 570, width: 760, height: 30 }, { fontSize: 20, bold: true, color: C.blue, alignment: "center" });
notes(chain, "Educational synthesis based on the standard chain rule in calculus.");

await clearBody(backprop);
await setHeader(backprop, "Learning", "Backpropagation applies the chain rule backward");
panel(backprop);
text(backprop, "The forward pass computes values; the backward pass computes how the loss changes each weight.", { left: 120, top: 180, width: 1040, height: 34 }, { fontSize: 22, color: C.muted, alignment: "center" });

const xs = [120, 390, 670, 960];
// Forward arrows and backward arrows first.
for (let i = 0; i < 3; i++) arrow(backprop, { left: xs[i] + 177, top: 278, width: 80, height: 38 });
for (let i = 0; i < 3; i++) arrow(backprop, { left: xs[i] + 177, top: 375, width: 80, height: 38 }, "left", C.red);
const labels = ["w", "z = wx + b", "a = σ(z)", "L(a)"];
const fills = [C.white, C.paleBlue, C.paleGreen, C.paleRed];
const lines = [C.line, C.blue, C.green, C.red];
for (let i = 0; i < 4; i++) box(backprop, labels[i], { left: xs[i], top: 250, width: 165, height: 90 }, { fontSize: i === 0 ? 28 : 23, fill: fills[i], line: lines[i], color: i === 3 ? C.red : C.ink });
text(backprop, "forward: compute activations and loss", { left: 355, top: 220, width: 570, height: 25 }, { fontSize: 17, color: C.blue, alignment: "center", bold: true });
text(backprop, "backward: propagate gradients toward earlier parameters", { left: 315, top: 420, width: 650, height: 25 }, { fontSize: 17, color: C.red, alignment: "center", bold: true });
box(backprop, "∂L/∂w  =  ∂L/∂a  ×  ∂a/∂z  ×  ∂z/∂w", { left: 165, top: 475, width: 950, height: 66 }, { fontSize: 27, fill: C.white, line: C.red, lineWidth: 1.8, color: C.ink });
box(backprop, "update:  w ← w − η · ∂L/∂w", { left: 350, top: 565, width: 580, height: 50 }, { fontSize: 22, fill: C.paleOrange, line: C.orange, color: "#C2410C" });
notes(backprop, "Educational synthesis based on standard backpropagation and gradient descent.");

// Re-number slide footers after insertion.
for (let i = 0; i < p.slides.items.length; i++) {
  const slide = p.slides.getItem(i);
  const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [x, y, w, h] = e.bbox ?? [];
    if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) {
      const footer = p.resolve(e.aid);
      footer.text = String(i + 1);
      footer.text.style = { typeface: "Calibri", fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

for (const n of [19, 20, 21, 22]) {
  const png = await p.export({ slide: p.slides.getItem(n - 1), format: "png", scale: 1 });
  await fs.writeFile(`${outDir}\\after-${n}.png`, new Uint8Array(await png.arrayBuffer()));
  const layout = await p.slides.getItem(n - 1).export({ format: "layout" });
  await fs.writeFile(`${outDir}\\after-${n}.layout.json`, await layout.text(), "utf8");
}
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
console.log(`slides=${p.slides.items.length}`);
