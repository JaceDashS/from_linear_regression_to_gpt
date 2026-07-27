import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\equations_verified.pptx";
const outDir = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2";
const p = await PresentationFile.importPptx(await FileBlob.load(source));

function svgDoc(width, height, body, color = "#0B0F19") {
  return `<?xml version="1.0" encoding="UTF-8"?>
  <svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
    <style>
      .m { font-family: "Cambria Math", "STIX Two Math", "Times New Roman", serif; font-size: 28px; font-style: italic; font-weight: 600; fill: ${color}; }
      .op { font-family: "Cambria Math", "Times New Roman", serif; font-size: 30px; font-weight: 600; fill: ${color}; }
      .small { font-family: "Cambria Math", "STIX Two Math", "Times New Roman", serif; font-size: 24px; font-style: italic; font-weight: 600; fill: ${color}; }
      .tiny { font-family: "Cambria Math", "STIX Two Math", "Times New Roman", serif; font-size: 20px; font-style: italic; font-weight: 600; fill: ${color}; }
      .bar { stroke: ${color}; stroke-width: 2; stroke-linecap: round; }
    </style>${body}</svg>`;
}
function fraction(cx, top, bottom, y = 6, half = 48, cls = "m") {
  return `<g><text x="${cx}" y="${y + 23}" text-anchor="middle" class="${cls}">${top}</text><line x1="${cx-half}" y1="${y+34}" x2="${cx+half}" y2="${y+34}" class="bar"/><text x="${cx}" y="${y+66}" text-anchor="middle" class="${cls}">${bottom}</text></g>`;
}
function compactFraction(cx, top, bottom, half = 34) {
  return `<g><text x="${cx}" y="17" text-anchor="middle" class="small">${top}</text><line x1="${cx-half}" y1="27" x2="${cx+half}" y2="27" class="bar"/><text x="${cx}" y="51" text-anchor="middle" class="small">${bottom}</text></g>`;
}
function tinyFraction(cx, top, bottom, half = 30) {
  return `<g><text x="${cx}" y="14" text-anchor="middle" class="tiny">${top}</text><line x1="${cx-half}" y1="23" x2="${cx+half}" y2="23" class="bar"/><text x="${cx}" y="44" text-anchor="middle" class="tiny">${bottom}</text></g>`;
}
function bytes(svg) { const b = Buffer.from(svg, "utf8"); return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength); }
function addSvg(slide, svg, position, alt) {
  return slide.images.add({ blob: bytes(svg), contentType: "image/svg+xml", alt, fit: "contain", position });
}

const chainMain = svgDoc(780, 78,
  `${fraction(78,"dy","dx",2,43)}<text x="160" y="48" class="op">=</text>${fraction(260,"dy","du",2,43)}<text x="344" y="48" class="op">·</text>${fraction(450,"du","dx",2,43)}`);
const chainExample = svgDoc(710, 58,
  `<text x="8" y="38" class="small">u = 3x + 1,   y = u</text><text x="185" y="22" class="small" font-size="16">2</text><line x1="240" y1="32" x2="288" y2="32" class="bar"/><polygon points="288,32 277,24 277,40" fill="#2563EB"/>${compactFraction(350,"dy","dx",32)}<text x="400" y="38" class="small">= 2u · 3</text>`, "#2563EB");
const backpropMain = svgDoc(900, 74,
  `${fraction(65,"∂L","∂w",0,40,"small")}<text x="132" y="45" class="op">=</text>${fraction(220,"∂L","∂a",0,40,"small")}<text x="292" y="45" class="op">·</text>${fraction(380,"∂a","∂z",0,40,"small")}<text x="452" y="45" class="op">·</text>${fraction(540,"∂z","∂w",0,40,"small")}`);
const updateEq = svgDoc(500, 58,
  `<text x="28" y="34" class="tiny">w ← w − η ·</text>${tinyFraction(225,"∂L","∂w",30)}`, "#C2410C");

const targets = {
  "sh/36tsfa98": { slide: 19, svg: chainMain, pos: { left: 250, top: 422, width: 780, height: 76 }, alt: "Chain rule equation with stacked derivative fractions" },
  "sh/kzqxs3a5": { slide: 19, svg: chainExample, pos: { left: 285, top: 558, width: 710, height: 58 }, alt: "Worked chain rule example rendered as an equation" },
  "sh/ah8byxw3": { slide: 20, svg: backpropMain, pos: { left: 210, top: 474, width: 900, height: 64 }, alt: "Backpropagation gradient rendered as a product of partial derivatives" },
  "sh/vihc7ido": { slide: 20, svg: updateEq, pos: { left: 430, top: 562, width: 500, height: 54 }, alt: "Gradient descent weight update equation" },
};
for (const [id, spec] of Object.entries(targets)) {
  const shape = p.resolve(id);
  shape.text = "";
  addSvg(p.slides.getItem(spec.slide), spec.svg, spec.pos, spec.alt);
}

for (const n of [20, 21]) {
  const slide = p.slides.getItem(n - 1);
  const png = await p.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(`${outDir}\\equation-slide-${n}.png`, new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outDir}\\equation-slide-${n}.layout.json`, await layout.text(), "utf8");
}
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
