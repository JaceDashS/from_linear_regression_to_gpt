import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const presentation = await PresentationFile.importPptx(
  await FileBlob.load("C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/candidate-v2.pptx"),
);
const slide = presentation.slides.getItem(53);
const png = await slide.export({ format: "png", scale: 1 });
await fs.mkdir("C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/verified-v2", { recursive: true });
await fs.writeFile(
  "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/verified-v2/slide-54.png",
  new Uint8Array(await png.arrayBuffer()),
);
await fs.writeFile(
  "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/verified-v2/slide-54.layout.json",
  await (await slide.export({ format: "layout" })).text(),
  "utf8",
);
console.log(JSON.stringify({ slides: presentation.slides.items.length }));
