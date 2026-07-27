import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/slide15-solid-candidate.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/slide15-solid-render";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = presentation.slides.getItem(14);

const dottedCurveLayer = slide.images.items.find((image) => {
  const box = image.position;
  return box && box.left === 0 && box.top === 0 && box.width === 1280 && box.height === 720;
});
if (!dottedCurveLayer) throw new Error("Dotted curve image layer not found");
dottedCurveLayer.delete();

const left = 184;
const top = 330;
const width = 762;
const height = 200;
const commands = [];
for (let x = 0; x <= width; x += 8) {
  const globalX = left + x;
  const globalY = 524 - 0.00132 * Math.pow(globalX - 565, 2);
  const point = { x, y: globalY - top };
  commands.push(x === 0 ? { moveTo: point } : { lineTo: point });
}

slide.shapes.add({
  geometry: "custom",
  name: "solid-loss-curve",
  position: { left, top, width, height },
  fill: "none",
  line: { style: "solid", fill: "#2563EB", width: 4 },
  customPaths: [{ width, height, commands }],
});

for (const shape of slide.shapes.items) {
  const text = shape.text?.toString().trim() ?? "";
  const box = shape.position;
  const isMarker = box && box.width >= 20 && box.width <= 30 && box.height >= 20 && box.height <= 30;
  const isAxis = box && ((box.width <= 2 && box.height >= 250) || (box.height <= 2 && box.width >= 700));
  if (isMarker || isAxis || ["reject", "accept", "MSE", "parameter", "α → α/2"].includes(text)) {
    shape.bringToFront();
  }
}

await fs.mkdir(OUT, { recursive: true });
await writeBlob(`${OUT}/slide-15.png`, await slide.export({ format: "png", scale: 2 }));
await fs.writeFile(`${OUT}/slide-15.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
