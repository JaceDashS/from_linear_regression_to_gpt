import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading.pptx";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\dotted_to_svg";
const SVG_PATH = `${TMP}\\assets\\slide14_dotted_curves.svg`;
const RENDER = `${TMP}\\final-render`;

function escapeXml(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const snapshot = await presentation.inspect({ kind: "shape,image", maxChars: 1500000 });
const dots = [];

for (const line of snapshot.ndjson.split(/\r?\n/)) {
  if (!line.trim()) continue;
  const item = JSON.parse(line);
  if (item.slide !== 14 || item.kind !== "shape" || !item.id) continue;
  const bbox = item.bbox ?? [];
  if (bbox.length !== 4 || bbox[2] > 10 || bbox[3] > 10) continue;
  const shape = presentation.resolve(item.id);
  const proto = shape.toProto?.();
  const geometry = proto?.shape?.geometry;
  const fillColor = proto?.shape?.fill?.color?.value;
  const lineColor = proto?.shape?.line?.fill?.color?.value;
  if (geometry !== 35 || !["2563EB", "EF4444", "10B981"].includes(fillColor) || lineColor !== fillColor) continue;
  dots.push({
    id: item.id,
    cx: bbox[0] + bbox[2] / 2,
    cy: bbox[1] + bbox[3] / 2,
    rx: bbox[2] / 2,
    ry: bbox[3] / 2,
    color: `#${fillColor}`,
  });
}

if (dots.length !== 115) throw new Error(`Expected 115 dotted-curve points on slide 14; found ${dots.length}.`);

const circles = dots.map((dot) =>
  `  <ellipse cx="${dot.cx.toFixed(2)}" cy="${dot.cy.toFixed(2)}" rx="${dot.rx.toFixed(2)}" ry="${dot.ry.toFixed(2)}" fill="${escapeXml(dot.color)}"/>`,
).join("\n");
const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="1280" height="720" viewBox="0 0 1280 720">
  <rect width="1280" height="720" fill="none"/>
${circles}
</svg>`;

await fs.mkdir(`${TMP}\\assets`, { recursive: true });
await fs.writeFile(SVG_PATH, svg, "utf8");

for (const dot of dots) presentation.resolve(dot.id).delete();

const bytes = await fs.readFile(SVG_PATH);
presentation.slides.getItem(13).images.add({
  blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
  contentType: "image/svg+xml",
  alt: "Backtracking loss curves with blue, red, and green dotted segments as SVG",
  fit: "contain",
  position: { left: 0, top: 0, width: 1280, height: 720 },
});

// The source dots sat below labels and decision markers. Restore that stacking
// relationship after adding the transparent full-slide SVG overlay.
for (const shape of presentation.slides.getItem(13).shapes.items) {
  const pos = shape.position;
  const proto = shape.toProto?.();
  const hasText = Boolean(shape.text?.toString().trim());
  const isLargeMarker = proto?.shape?.geometry === 35 && pos?.width >= 20 && pos?.height >= 20;
  if (hasText || isLargeMarker) shape.bringToFront();
}

const frameMap = {
  outputSlides: presentation.slides.items.map((_, index) => ({
    outputSlide: index + 1,
    sourceSlide: index + 1,
    narrativeRole: index === 13 ? "backtracking dotted loss curves" : "preserve source slide",
    reuseMode: "duplicate-slide",
    editTargets: index === 13 ? dots.map((dot) => ({ sourceElementId: dot.id, action: "replace", replacement: "single embedded SVG" })) : [],
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
console.log(`replaced=${dots.length}`);
console.log(`saved=${FINAL}`);
