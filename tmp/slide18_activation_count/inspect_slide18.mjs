import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const p=await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
const slide=p.slides.getItem(17);
const png=await p.export({slide,format:"png",scale:1});
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\slide-18-before.png",Buffer.from(await png.arrayBuffer()));
const layout=await slide.export({format:"layout"});
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\slide-18-before.layout.json",await layout.text());
