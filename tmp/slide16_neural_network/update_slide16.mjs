import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide16_neural_network\\edited.pptx";
const preview = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide16_neural_network\\slide-16-after.png";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));

presentation.resolve("sh/i1krq9wz").text.replace("Depth", "Neural networks");
presentation.resolve("sh/lcfqx4ji").text.replace(
  "A network stacks many neurons",
  "Neural network: basic structure",
);

const slide = presentation.slides.getItem(15);
const png = await presentation.export({ slide, format: "png", scale: 1 });
await fs.writeFile(preview, Buffer.from(await png.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
