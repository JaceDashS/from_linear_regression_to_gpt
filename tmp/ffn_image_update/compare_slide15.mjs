import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const files = [
  ["before", "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/candidate-v2.pptx"],
  ["current", "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx"],
];
const out = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/slide15-compare";
await fs.mkdir(out, { recursive: true });

for (const [label, path] of files) {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(path));
  const slide = presentation.slides.getItem(14);
  const png = await slide.export({ format: "png", scale: 1 });
  await fs.writeFile(`${out}/${label}.png`, new Uint8Array(await png.arrayBuffer()));
  await fs.writeFile(`${out}/${label}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
  const texts = slide.shapes.items.map((shape) => shape.text?.toString().trim()).filter(Boolean);
  console.log(JSON.stringify({ label, slides: presentation.slides.items.length, texts }));
}
