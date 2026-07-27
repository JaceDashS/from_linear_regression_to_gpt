import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const presentation = await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
for (const index of [18,19]) {
  const slide = presentation.slides.getItem(index);
  const png = await presentation.export({slide,format:"png",scale:1});
  await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\activation_image\\slide-${index+1}-review.png`,Buffer.from(await png.arrayBuffer()));
  const layout=JSON.parse(await (await slide.export({format:"layout"})).text());
  console.log(JSON.stringify({slide:index+1,text:layout.elements.filter(x=>x.text).map(x=>x.text)}));
}
