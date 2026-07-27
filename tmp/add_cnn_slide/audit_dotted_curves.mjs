import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added.pptx";
const OUT = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_cnn_slide\\dotted-curve-audit.txt";

const presentation = await PresentationFile.importPptx(await FileBlob.load(FINAL));
const shapes = await presentation.inspect({ kind: "shape", maxChars: 1500000 });
const images = await presentation.inspect({ kind: "image", maxChars: 300000 });
const candidateCounts = new Map();

for (const line of shapes.ndjson.split(/\r?\n/)) {
  if (!line.trim()) continue;
  const item = JSON.parse(line);
  const bbox = item.bbox ?? [];
  if (bbox.length !== 4 || bbox[2] > 10 || bbox[3] > 10 || !item.id) continue;
  const proto = presentation.resolve(item.id).toProto?.();
  const geometry = proto?.shape?.geometry;
  const fillColor = proto?.shape?.fill?.color?.value;
  const lineColor = proto?.shape?.line?.fill?.color?.value;
  if (geometry === 35 && fillColor && lineColor === fillColor) {
    const key = `${item.slide}:${fillColor}`;
    candidateCounts.set(key, (candidateCounts.get(key) ?? 0) + 1);
  }
}

const svgImages = [];
for (const line of images.ndjson.split(/\r?\n/)) {
  if (!line.trim()) continue;
  const item = JSON.parse(line);
  if (/dotted|svg/i.test(`${item.alt ?? ""} ${item.name ?? ""}`)) svgImages.push(item);
}

const clusters = [...candidateCounts.entries()].filter(([, count]) => count >= 12);
const report = [
  `slides=${presentation.slides.items.length}`,
  `remainingDottedPointClusters=${clusters.length}`,
  ...clusters.map(([key, count]) => `cluster=${key},count=${count}`),
  `svgDottedCurveImages=${svgImages.length}`,
  ...svgImages.map((item) => `svgImageSlide=${item.slide},id=${item.id},alt=${item.alt ?? ""}`),
].join("\n") + "\n";
await fs.writeFile(OUT, report, "utf8");
console.log(report);
