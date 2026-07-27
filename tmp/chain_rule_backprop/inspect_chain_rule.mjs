import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const presentation=await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
const inspection=await presentation.inspect({kind:"slide,textbox,shape,notes",search:"chain",maxChars:30000});
console.log(inspection.ndjson);
for (const index of [20,21,22,23]) {
  const slide=presentation.slides.getItem(index);
  const png=await presentation.export({slide,format:"png",scale:1});
  const fs=await import("node:fs/promises");
  await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop\\slide-${index+1}-before.png`,Buffer.from(await png.arrayBuffer()));
}
