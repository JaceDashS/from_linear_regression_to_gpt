import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
for (const index of [16,17]) {
  const slide = presentation.slides.getItem(index);
  const layout = await slide.export({format:"layout"});
  await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\bias_perceptron\\slide-${index+1}.layout.json`,await layout.text());
}
