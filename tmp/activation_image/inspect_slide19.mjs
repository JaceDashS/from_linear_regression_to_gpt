import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const slide = presentation.slides.getItem(18);
const png = await presentation.export({slide,format:"png",scale:1});
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\activation_image\\slide-19-before.png",Buffer.from(await png.arrayBuffer()));
const layout = await slide.export({format:"layout"});
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\activation_image\\slide-19-before.layout.json",await layout.text());
