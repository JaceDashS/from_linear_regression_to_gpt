import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const p=await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
const inspection=await p.inspect({kind:"slide,textbox,shape,image,notes",search:"CNNs reuse filters",maxChars:10000});
console.log(inspection.ndjson);
for(let i=0;i<p.slides.items.length;i++){const slide=p.slides.getItem(i);const layout=JSON.parse(await (await slide.export({format:"layout"})).text());if(layout.elements.some(e=>e.text==="CNNs reuse filters to find local patterns")){const png=await p.export({slide,format:"png",scale:1});await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_repair\\cnn-before.png",Buffer.from(await png.arrayBuffer()));console.log(`INDEX=${i}`);}}
