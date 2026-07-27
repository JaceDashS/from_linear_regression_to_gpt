import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\goal_slide1\\edited.pptx";
const preview = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\goal_slide1\\slide-1.png";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const goal = presentation.resolve("sh/jetofylo");

goal.text.replace(
  "Goal: understand why attention made GPT possible",
  "Goal: understand GPT's core principles and architecture\nto assess engineering possibilities",
);

const slide = presentation.slides.getItem(0);
const png = await presentation.export({ slide, format: "png", scale: 1 });
await fs.writeFile(preview, Buffer.from(await png.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
console.log(output);
