import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const input = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves.pptx";
const out = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\dotted_to_svg\\verified";
await fs.mkdir(out, { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(input));
const snapshot = await presentation.inspect({ kind: "shape,image", maxChars: 1500000 });
let remainingCurveDots = 0;
let svgOverlay = 0;

for (const line of snapshot.ndjson.split(/\r?\n/)) {
  if (!line.trim()) continue;
  const item = JSON.parse(line);
  if (item.slide !== 14 || !item.id) continue;
  const bbox = item.bbox ?? [];
  if (item.kind === "image" && bbox.length === 4 && bbox[0] === 0 && bbox[1] === 0 && bbox[2] === 1280 && bbox[3] === 720) svgOverlay += 1;
  if (item.kind !== "shape") continue;
  const proto = presentation.resolve(item.id).toProto?.();
  const fill = proto?.shape?.fill?.color?.value;
  if (proto?.shape?.geometry === 35 && bbox.length === 4 && bbox[2] <= 10 && bbox[3] <= 10 && ["2563EB", "EF4444", "10B981"].includes(fill)) {
    remainingCurveDots += 1;
  }
}

const slide = presentation.slides.getItem(13);
const png = await slide.export({ format: "png", scale: 2 });
await fs.writeFile(`${out}\\slide-14.png`, new Uint8Array(await png.arrayBuffer()));
const layout = await slide.export({ format: "layout" });
await fs.writeFile(`${out}\\slide-14.layout.json`, await layout.text(), "utf8");
await fs.writeFile(`${out}\\verification.txt`, [
  `slides=${presentation.slides.items.length}`,
  `remainingCurveDots=${remainingCurveDots}`,
  `svgOverlay=${svgOverlay}`,
].join("\n"), "utf8");

if (presentation.slides.items.length !== 52 || remainingCurveDots !== 0 || svgOverlay !== 1) {
  throw new Error(`Verification failed: slides=${presentation.slides.items.length}, remaining=${remainingCurveDots}, svg=${svgOverlay}`);
}
console.log(`slides=${presentation.slides.items.length}`);
console.log(`remainingCurveDots=${remainingCurveDots}`);
console.log(`svgOverlay=${svgOverlay}`);
