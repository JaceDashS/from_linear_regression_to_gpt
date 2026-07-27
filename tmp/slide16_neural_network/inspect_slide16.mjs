import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const slide = presentation.slides.getItem(15);
const png = await presentation.export({ slide, format: "png", scale: 1 });
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide16_neural_network\\slide-16-before.png", Buffer.from(await png.arrayBuffer()));
const layout = await slide.export({ format: "layout" });
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide16_neural_network\\slide-16-before.layout.json", await layout.text());
const inspection = await presentation.inspect({ kind: "slide,textbox,shape,notes", maxChars: 30000 });
for (const row of inspection.ndjson.split(/\r?\n/)) {
  if (!row.trim()) continue;
  const item = JSON.parse(row);
  if (item.slide !== 16) continue;
  const object = presentation.resolve(item.id);
  console.log(JSON.stringify({ ...item, text: object?.text?.text ?? null }));
}
