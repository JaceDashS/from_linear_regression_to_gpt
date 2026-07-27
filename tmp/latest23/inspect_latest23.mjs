import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const p=await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
const s=p.slides.getItem(22);const b=await p.export({slide:s,format:"png",scale:1});await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\latest23\\slide23.png",new Uint8Array(await b.arrayBuffer()));
const l=await (await s.export({format:"layout"})).text();await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\latest23\\slide23.layout.json",l);console.log(l);
