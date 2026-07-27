import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const input = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading.pptx";
const out = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\dotted_to_svg";

const presentation = await PresentationFile.importPptx(await FileBlob.load(input));
const snapshot = await presentation.inspect({ kind: "slide,shape,image", maxChars: 1500000 });
const rows = [];
const tiny = [];
const tinyBySlide = new Map();

for (const line of snapshot.ndjson.split(/\r?\n/)) {
  if (!line.trim()) continue;
  const item = JSON.parse(line);
  if (!item.id || !["shape", "image"].includes(item.kind)) continue;
  const resolved = presentation.resolve(item.id);
  const proto = resolved.toProto?.();
  const bbox = item.bbox ?? [];
  const shape = proto?.shape;
  const record = {
    slide: item.slide,
    id: item.id,
    kind: item.kind,
    bbox,
    geometry: shape?.geometry,
    lineStyle: shape?.line?.style,
    lineColor: shape?.line?.fill?.color?.value,
    lineWidth: shape?.line?.width,
    fillColor: shape?.fill?.color?.value,
    contentType: item.contentType,
    alt: item.alt,
  };
  if (item.kind === "image" && item.contentType === "image/svg+xml") rows.push({ category: "svg", ...record });
  if (shape?.line?.style && shape.line.style !== 1) rows.push({ category: "non-solid-line", ...record });
  if (bbox.length === 4 && bbox[2] <= 10 && bbox[3] <= 10 && shape?.geometry) {
    const count = tinyBySlide.get(item.slide) ?? 0;
    tinyBySlide.set(item.slide, count + 1);
    tiny.push(record);
  }
}

await fs.mkdir(out, { recursive: true });
await fs.writeFile(`${out}\\dotted-audit.json`, JSON.stringify({ rows, tiny, tinyBySlide: Object.fromEntries(tinyBySlide) }, null, 2));
console.log(JSON.stringify({ slides: presentation.slides.items.length, records: rows.length, tinyBySlide: Object.fromEntries(tinyBySlide) }, null, 2));
