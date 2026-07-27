import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const input = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading.pptx";
const out = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\llm_load_slide\\verified";
await fs.mkdir(out, { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(input));
const report = await presentation.inspect({
  kind: "slide,textbox,shape,notes,layout",
  search: "An LLM needs three matching parts",
  maxChars: 12000,
});
await fs.writeFile(`${out}\\focused-inspect.ndjson`, report.ndjson, "utf8");

for (const slideNumber of [1, 40, 41, 42, 52]) {
  const slide = presentation.slides.getItem(slideNumber - 1);
  const png = await slide.export({ format: "png", scale: 2 });
  await fs.writeFile(`${out}\\slide-${slideNumber}.png`, new Uint8Array(await png.arrayBuffer()));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${out}\\slide-${slideNumber}.layout.json`, await layout.text(), "utf8");
}

console.log(`slides=${presentation.slides.items.length}`);
