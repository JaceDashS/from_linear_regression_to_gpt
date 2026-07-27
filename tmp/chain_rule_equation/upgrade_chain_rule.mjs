import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves.pptx";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\chain_rule_equation";
const SVG_PATH = `${TMP}\\assets\\slide15_chain_rule.svg`;
const RENDER = `${TMP}\\final-render`;

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function fraction(cx, numerator, denominator) {
  return `
  <g aria-label="${numerator} over ${denominator}">
    <text x="${cx}" y="33" text-anchor="middle" class="math">${numerator}</text>
    <line x1="${cx - 52}" y1="44" x2="${cx + 52}" y2="44" class="bar"/>
    <text x="${cx}" y="82" text-anchor="middle" class="math">${denominator}</text>
  </g>`;
}

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="980" height="96" viewBox="0 0 980 96" role="img" aria-label="Partial L over partial w equals partial L over partial p times partial p over partial a times partial a over partial z times partial z over partial w">
  <style>
    .math { font-family: "Cambria Math", "STIX Two Math", "Times New Roman", serif; font-size: 34px; font-style: italic; font-weight: 600; fill: #2563EB; }
    .op { font-family: "Cambria Math", "Times New Roman", serif; font-size: 36px; font-weight: 600; fill: #2563EB; }
    .bar { stroke: #2563EB; stroke-width: 2.5; stroke-linecap: round; }
  </style>
  ${fraction(80, "∂L", "∂w")}
  <text x="185" y="58" text-anchor="middle" class="op">=</text>
  ${fraction(290, "∂L", "∂p")}
  <text x="395" y="58" text-anchor="middle" class="op">·</text>
  ${fraction(500, "∂p", "∂a")}
  <text x="605" y="58" text-anchor="middle" class="op">·</text>
  ${fraction(710, "∂a", "∂z")}
  <text x="815" y="58" text-anchor="middle" class="op">·</text>
  ${fraction(920, "∂z", "∂w")}
</svg>`;

await fs.mkdir(`${TMP}\\assets`, { recursive: true });
await fs.writeFile(SVG_PATH, svg, "utf8");

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const snapshot = await presentation.inspect({ kind: "textbox,shape,image", search: "dL/dw", maxChars: 12000 });
let formulaId;
for (const line of snapshot.ndjson.split(/\r?\n/)) {
  if (!line.trim()) continue;
  const item = JSON.parse(line);
  if (item.slide === 15 && item.kind === "textbox" && String(item.text ?? "").includes("dL/dw")) formulaId = item.id;
}
if (!formulaId) throw new Error("Could not locate the existing chain-rule formula on slide 15.");
presentation.resolve(formulaId).delete();

const bytes = await fs.readFile(SVG_PATH);
presentation.slides.getItem(14).images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/svg+xml",
  alt: "Chain rule rendered as five stacked partial-derivative fractions",
  fit: "contain",
  position: { left: 150, top: 474, width: 980, height: 96 },
});

const frameMap = {
  outputSlides: presentation.slides.items.map((_, index) => ({
    outputSlide: index + 1,
    sourceSlide: index + 1,
    narrativeRole: index === 14 ? "chain-rule equation" : "preserve source slide",
    reuseMode: "duplicate-slide",
    editTargets: index === 14 ? [{ sourceElementId: formulaId, action: "replace", replacement: "stacked-fraction SVG equation" }] : [],
  })),
  omittedSourceSlides: [],
};
await fs.writeFile(`${TMP}\\template-frame-map.json`, JSON.stringify(frameMap, null, 2), "utf8");

await fs.mkdir(`${RENDER}\\layouts`, { recursive: true });
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const slide = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${RENDER}\\${stem}.png`, await slide.export({ format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${RENDER}\\layouts\\${stem}.layout.json`, await layout.text(), "utf8");
}
await writeBlob(`${RENDER}\\montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(FINAL);
console.log(`replaced=${formulaId}`);
console.log(`saved=${FINAL}`);
