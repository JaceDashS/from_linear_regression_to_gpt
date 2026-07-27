import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\u_math_verified.pptx";
const outDir = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const slide = p.slides.getItem(20);
const C = { ink: "#0B0F19", muted: "#667085", line: "#CBD5E1", blue: "#2563EB", paleBlue: "#EFF6FF", green: "#059669", paleGreen: "#ECFDF5", red: "#EF4444", paleRed: "#FEF2F2", orange: "#F97316", paleOrange: "#FFF7ED", white: "#FFFFFF", panel: "#F8FAFC" };

function text(value, position, style = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  s.text = value;
  s.text.style = { typeface: "Calibri", fontSize: style.fontSize ?? 18, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  if (style.verticalAlignment) s.text.verticalAlignment = style.verticalAlignment;
  return s;
}
function box(position, opts = {}) { return slide.shapes.add({ geometry: opts.geometry ?? "roundRect", position, fill: opts.fill ?? C.white, line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.4 } }); }
function arrow(position, color = C.blue) { return box(position, { geometry: "rightArrow", fill: color === C.blue ? C.paleBlue : C.paleGreen, line: color }); }
function svg(width, height, body, color = C.ink) { return `<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><style>.m{font-family:"Cambria Math","Times New Roman",serif;font-size:30px;font-style:italic;font-weight:600;fill:${color}}.s{font-family:"Cambria Math","Times New Roman",serif;font-size:23px;font-style:italic;font-weight:600;fill:${color}}.op{font-family:"Cambria Math","Times New Roman",serif;font-size:27px;font-weight:600;fill:${color}}.bar{stroke:${color};stroke-width:2;stroke-linecap:round}</style>${body}</svg>`; }
function frac(cx, top, bottom, half = 35) { return `<g><text x="${cx}" y="18" text-anchor="middle" class="s">${top}</text><line x1="${cx-half}" y1="28" x2="${cx+half}" y2="28" class="bar"/><text x="${cx}" y="52" text-anchor="middle" class="s">${bottom}</text></g>`; }
function addSvg(svgText, position, alt) { const b = Buffer.from(svgText, "utf8"); slide.images.add({ blob: b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength), contentType: "image/svg+xml", alt, fit: "contain", position }); }

// Clear only the slide body; retain inherited background, header, and footer.
let layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) { const [, y] = e.bbox ?? []; if (e.aid && y >= 142 && y <= 642) p.resolve(e.aid).delete(); }
for (const image of [...slide.images.items]) image.delete();

layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) {
  const [x, y] = e.bbox ?? [];
  if (!e.aid || x < 60 || x > 80) continue;
  const s = p.resolve(e.aid);
  if (y >= 30 && y <= 40) { s.text = "Example"; s.text.style = { typeface: "Georgia", fontSize: 16, bold: true, color: C.blue, alignment: "left" }; s.text.verticalAlignment = "middle"; }
  if (y >= 58 && y <= 70) { s.text = "u splits the calculation into two steps"; s.text.style = { typeface: "Georgia", fontSize: 38, bold: true, color: C.ink, alignment: "left" }; s.text.verticalAlignment = "middle"; }
}

box({ left: 64, top: 142, width: 1152, height: 500 }, { fill: C.panel, line: "#E2E8F0", lineWidth: 1 });

text("1   Define", { left: 120, top: 170, width: 180, height: 26 }, { fontSize: 18, bold: true, color: C.blue });
arrow({ left: 585, top: 216, width: 90, height: 40 });
box({ left: 175, top: 204, width: 360, height: 66 }, { fill: C.paleBlue, line: C.blue });
box({ left: 725, top: 204, width: 360, height: 66 }, { fill: C.paleGreen, line: C.green });
addSvg(svg(320, 52, `<text x="66" y="38" class="m">u = 3x + 1</text>`, C.blue), { left: 195, top: 211, width: 320, height: 50 }, "Equation defining intermediate variable u");
addSvg(svg(320, 52, `<text x="92" y="38" class="m">y = u</text><text x="170" y="22" class="s">2</text>`, C.green), { left: 745, top: 211, width: 320, height: 50 }, "Equation defining y as u squared");

text("2   Evaluate at x = 2", { left: 120, top: 298, width: 260, height: 26 }, { fontSize: 18, bold: true, color: C.green });
arrow({ left: 585, top: 343, width: 90, height: 40 }, C.green);
box({ left: 175, top: 331, width: 360, height: 66 }, { fill: C.white, line: C.line });
box({ left: 725, top: 331, width: 360, height: 66 }, { fill: C.white, line: C.line });
addSvg(svg(320, 52, `<text x="38" y="38" class="m">u = 3(2) + 1 = 7</text>`, C.ink), { left: 195, top: 338, width: 320, height: 50 }, "Substitution x equals 2 gives u equals 7");
addSvg(svg(320, 52, `<text x="67" y="38" class="m">y = 7</text><text x="135" y="22" class="s">2</text><text x="163" y="38" class="m">= 49</text>`, C.ink), { left: 745, top: 338, width: 320, height: 50 }, "Substitution u equals 7 gives y equals 49");

text("3   Differentiate", { left: 120, top: 426, width: 220, height: 26 }, { fontSize: 18, bold: true, color: C.red });
box({ left: 150, top: 463, width: 980, height: 70 }, { fill: C.white, line: C.red, lineWidth: 1.7 });
const chainEq = svg(900, 58, `${frac(68,"dy","dx")}<text x="132" y="40" class="op">=</text>${frac(210,"dy","du")}<text x="274" y="40" class="op">·</text>${frac(350,"du","dx")}<text x="414" y="40" class="op">= 2u · 3</text>`, C.ink);
addSvg(chainEq, { left: 190, top: 469, width: 900, height: 58 }, "Chain rule derivative through intermediate variable u");

box({ left: 330, top: 558, width: 620, height: 52 }, { fill: C.paleOrange, line: C.orange, lineWidth: 1.5 });
const resultEq = svg(560, 46, `<text x="28" y="34" class="s">u = 7</text><line x1="130" y1="25" x2="176" y2="25" class="bar"/><polygon points="176,25 165,17 165,33" fill="#F97316"/><text x="210" y="34" class="s">dy/dx = 2 · 7 · 3 = 42</text>`, "#C2410C");
addSvg(resultEq, { left: 360, top: 561, width: 560, height: 46 }, "Final derivative at x equals 2 is 42");

slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on the standard chain rule in calculus.");
slide.speakerNotes.setVisible(true);

for (const n of [20, 21, 22]) { const target = p.slides.getItem(n - 1); const png = await p.export({ slide: target, format: "png", scale: 1 }); await fs.writeFile(`${outDir}\\u-math-${n}.png`, new Uint8Array(await png.arrayBuffer())); const l = await target.export({ format: "layout" }); await fs.writeFile(`${outDir}\\u-math-${n}.layout.json`, await l.text(), "utf8"); }
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
