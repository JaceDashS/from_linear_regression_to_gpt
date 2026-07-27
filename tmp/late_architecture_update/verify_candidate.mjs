import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const presentation = await PresentationFile.importPptx(
  await FileBlob.load("C:/docs/from_linear_regression_to_gpt/tmp/late_architecture_update/candidate.pptx"),
);
const out = "C:/docs/from_linear_regression_to_gpt/tmp/late_architecture_update/verified";
await fs.mkdir(out, { recursive: true });
for (const slideNumber of [55, 56]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  const png = await slide.export({ format: "png", scale: 1 });
  await fs.writeFile(`${out}/slide-${slideNumber}.png`, new Uint8Array(await png.arrayBuffer()));
  await fs.writeFile(`${out}/slide-${slideNumber}.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
}
console.log(JSON.stringify({ slides: presentation.slides.items.length }));
