import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/depth2-normalized.pptx";
const out = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/depth2-verified";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
if (presentation.slides.items.length !== 59) throw new Error("Slide count changed");
await fs.mkdir(out, { recursive: true });
for (let i = 15; i <= 17; i += 1) {
  const slide = presentation.slides.getItem(i);
  const png = await slide.export({ format: "png", scale: 1 });
  await fs.writeFile(`${out}/slide-${i + 1}.png`, new Uint8Array(await png.arrayBuffer()));
  await fs.writeFile(`${out}/slide-${i + 1}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}
console.log(JSON.stringify({ slides: presentation.slides.items.length }));
