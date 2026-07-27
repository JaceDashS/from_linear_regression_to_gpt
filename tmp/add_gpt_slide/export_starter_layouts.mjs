import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";
const input = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_gpt_slide\\template-starter.pptx";
const out = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_gpt_slide\\template-starter-layout";
await fs.mkdir(out, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(input));
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}.layout.json`;
  await fs.writeFile(`${out}\\${stem}`, await (await presentation.slides.getItem(index).export({ format: "layout" })).text(), "utf8");
}
console.log(`layouts=${presentation.slides.items.length}`);
