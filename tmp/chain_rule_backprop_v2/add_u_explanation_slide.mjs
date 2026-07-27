import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\u_explanation_verified.pptx";
const outDir = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const C = { ink: "#0B0F19", muted: "#667085", line: "#CBD5E1", blue: "#2563EB", paleBlue: "#EFF6FF", green: "#059669", paleGreen: "#ECFDF5", orange: "#F97316", paleOrange: "#FFF7ED", white: "#FFFFFF", panel: "#F8FAFC" };

function text(slide, value, position, style = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  s.text = value;
  s.text.style = { typeface: "Calibri", fontSize: style.fontSize ?? 20, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  if (style.verticalAlignment) s.text.verticalAlignment = style.verticalAlignment;
  return s;
}
function box(slide, value, position, opts = {}) {
  const s = slide.shapes.add({ geometry: opts.geometry ?? "roundRect", position, fill: opts.fill ?? C.white, line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 } });
  if (value) { s.text = value; s.text.style = { typeface: "Calibri", fontSize: opts.fontSize ?? 23, bold: opts.bold ?? true, color: opts.color ?? C.ink, alignment: opts.alignment ?? "center" }; s.text.verticalAlignment = "middle"; }
  return s;
}
function arrow(slide, position) { return box(slide, "", position, { geometry: "rightArrow", fill: C.paleBlue, line: C.blue, lineWidth: 1.4 }); }
function mathSvg(width, height, body, color = "#0B0F19") { return `<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><style>.m{font-family:"Cambria Math","Times New Roman",serif;font-size:24px;font-style:italic;font-weight:600;fill:${color}}.op{font-family:"Cambria Math","Times New Roman",serif;font-size:26px;font-weight:600;fill:${color}}.bar{stroke:${color};stroke-width:2;stroke-linecap:round}</style>${body}</svg>`; }
function frac(cx, top, bottom, half = 33) { return `<g><text x="${cx}" y="17" text-anchor="middle" class="m">${top}</text><line x1="${cx-half}" y1="27" x2="${cx+half}" y2="27" class="bar"/><text x="${cx}" y="51" text-anchor="middle" class="m">${bottom}</text></g>`; }
function addSvg(slide, svg, position, alt) { const b = Buffer.from(svg, "utf8"); slide.images.add({ blob: b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength), contentType: "image/svg+xml", alt, fit: "contain", position }); }

const chain = p.slides.getItem(19); // page 20, after Common activation functions
const uSlide = chain.duplicate();

// Move the small worked example out of page 20; the dedicated page below explains it instead.
let layout = JSON.parse(await (await chain.export({ format: "layout" })).text());
for (const e of layout.elements) { const [, y] = e.bbox ?? []; if (e.aid && y >= 550 && y <= 630) p.resolve(e.aid).delete(); }

// Rebuild the duplicate while preserving its inherited background, header structure, and footer.
layout = JSON.parse(await (await uSlide.export({ format: "layout" })).text());
for (const e of layout.elements) { const [, y] = e.bbox ?? []; if (e.aid && y >= 142 && y <= 642) p.resolve(e.aid).delete(); }
for (const image of [...uSlide.images.items]) image.delete();
layout = JSON.parse(await (await uSlide.export({ format: "layout" })).text());
for (const e of layout.elements) {
  const [x, y] = e.bbox ?? [];
  if (!e.aid || x < 60 || x > 80) continue;
  const s = p.resolve(e.aid);
  if (y >= 30 && y <= 40) { s.text = "Example"; s.text.style = { typeface: "Georgia", fontSize: 16, bold: true, color: C.blue, alignment: "left" }; s.text.verticalAlignment = "middle"; }
  if (y >= 58 && y <= 70) { s.text = "u names the intermediate result"; s.text.style = { typeface: "Georgia", fontSize: 38, bold: true, color: C.ink, alignment: "left" }; s.text.verticalAlignment = "middle"; }
}

box(uSlide, "", { left: 64, top: 142, width: 1152, height: 500 }, { fill: C.panel, line: "#E2E8F0", lineWidth: 1 });
text(uSlide, "u is not a new input—it simply stores the output of the first step.", { left: 150, top: 174, width: 980, height: 34 }, { fontSize: 22, color: C.muted, alignment: "center" });

// Forward computation arrows are added before the step boxes.
arrow(uSlide, { left: 312, top: 270, width: 80, height: 42 });
arrow(uSlide, { left: 588, top: 270, width: 80, height: 42 });
arrow(uSlide, { left: 864, top: 270, width: 80, height: 42 });
box(uSlide, "x = 2", { left: 132, top: 245, width: 165, height: 90 }, { fontSize: 28, fill: C.white, line: C.line });
box(uSlide, "u = 3x + 1", { left: 408, top: 245, width: 165, height: 90 }, { fontSize: 25, fill: C.paleBlue, line: C.blue, color: C.blue });
box(uSlide, "u = 7", { left: 684, top: 245, width: 165, height: 90 }, { fontSize: 28, fill: C.paleGreen, line: C.green, color: C.green });
box(uSlide, "y = u² = 49", { left: 960, top: 245, width: 165, height: 90 }, { fontSize: 24, fill: C.paleOrange, line: C.orange, color: "#C2410C" });
text(uSlide, "first operation", { left: 402, top: 345, width: 175, height: 24 }, { fontSize: 16, color: C.blue, alignment: "center", bold: true });
text(uSlide, "intermediate value", { left: 675, top: 345, width: 185, height: 24 }, { fontSize: 16, color: C.green, alignment: "center", bold: true });
text(uSlide, "second operation", { left: 953, top: 345, width: 180, height: 24 }, { fontSize: 16, color: "#C2410C", alignment: "center", bold: true });

box(uSlide, "Because y depends on x through u, the chain rule connects the two steps.", { left: 160, top: 410, width: 960, height: 58 }, { fontSize: 21, fill: C.white, line: C.line, color: C.ink });
const derivative = mathSvg(760, 60, `${frac(65,"dy","dx")}<text x="128" y="40" class="op">=</text>${frac(205,"dy","du")}<text x="268" y="40" class="op">·</text>${frac(345,"du","dx")}<text x="413" y="40" class="op">= 2u · 3 = 42</text>`, C.blue);
addSvg(uSlide, derivative, { left: 260, top: 505, width: 760, height: 58 }, "Worked chain rule derivative with intermediate variable u");
text(uSlide, "At x = 2, u = 7; so the local slope 2u becomes 14, and 14 × 3 = 42.", { left: 175, top: 578, width: 930, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });
uSlide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on the standard chain rule in calculus.");
uSlide.speakerNotes.setVisible(true);

for (let i = 0; i < p.slides.items.length; i++) {
  const slide = p.slides.getItem(i);
  const list = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of list.elements) { const [x, y, w, h] = e.bbox ?? []; if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) { const footer = p.resolve(e.aid); footer.text = String(i + 1); footer.text.style = { typeface: "Calibri", fontSize: 13, color: C.muted, alignment: "right" }; } }
}

for (const n of [20, 21, 22, 23]) { const png = await p.export({ slide: p.slides.getItem(n - 1), format: "png", scale: 1 }); await fs.writeFile(`${outDir}\\u-explanation-${n}.png`, new Uint8Array(await png.arrayBuffer())); }
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
console.log(`slides=${p.slides.items.length}`);
