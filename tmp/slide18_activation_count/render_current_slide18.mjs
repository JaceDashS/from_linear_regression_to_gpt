import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const p=await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
const png=await p.export({slide:p.slides.getItem(17),format:"png",scale:1});
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\slide-18-current.png",Buffer.from(await png.arrayBuffer()));
