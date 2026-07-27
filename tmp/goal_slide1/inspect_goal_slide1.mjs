import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const inspection = await presentation.inspect({
  kind: "slide,textbox,shape,notes",
  maxChars: 30000,
});

for (const row of inspection.ndjson.split(/\r?\n/)) {
  if (!row.trim()) continue;
  const item = JSON.parse(row);
  if (item.slide === 1) {
    const object = presentation.resolve(item.id);
    console.log(JSON.stringify({ ...item, text: object?.text?.text ?? object?.text?.toString?.() ?? null }));
  }
}
