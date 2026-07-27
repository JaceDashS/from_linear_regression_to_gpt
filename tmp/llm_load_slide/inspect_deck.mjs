import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const outDir = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\llm_load_slide";

const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const inventory = await presentation.inspect({
  kind: "slide,textbox,shape,notes,layout",
  maxChars: 120000,
});
await fs.writeFile(`${outDir}\\deck-inspect.ndjson`, inventory.ndjson, "utf8");

for (const slideNumber of [39, 40, 41, 42]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  const png = await slide.export({ format: "png", scale: 2 });
  await fs.writeFile(`${outDir}\\source-slide-${slideNumber}.png`, new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outDir}\\source-slide-${slideNumber}.layout.json`, await layout.text(), "utf8");
}

console.log(`slides=${presentation.slides.items.length}`);
