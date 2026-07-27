import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\bias_perceptron\\edited.pptx";
const outDir = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\bias_perceptron";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));

async function elements(slide) {
  return JSON.parse(await (await slide.export({format:"layout"})).text()).elements;
}
function replaceByText(items, oldText, newText) {
  const item = items.find(entry => entry.text === oldText);
  if (!item) throw new Error(`Could not find: ${oldText}`);
  presentation.resolve(item.aid).text.replace(oldText,newText);
}
function addText(slide, text, position, color) {
  const shape = slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}});
  shape.text=text; shape.text.style={fontSize:16,bold:true,color,alignment:"center"};
}

const network = presentation.slides.getItem(16);
const networkItems = await elements(network);
replaceByText(networkItems,"w₁, b₁","w₁");
replaceByText(networkItems,"w₂, b₂","w₂");
replaceByText(networkItems,"w₃, b₃","w₃");
addText(network,"+ b₁",{left:418,top:382,width:84,height:22},"#2563EB");
addText(network,"+ b₂",{left:658,top:382,width:84,height:22},"#7C3AED");
addText(network,"+ b₃",{left:973,top:382,width:84,height:22},"#10B981");

const parameters = presentation.slides.getItem(17);
const parameterItems = await elements(parameters);
replaceByText(parameterItems,"z = wx + b","perceptron\nz = wx + b");
replaceByText(parameterItems,"bias b","bias b");
replaceByText(parameterItems,"b shifts the threshold","each perceptron\nhas its own bias");

for (const [index,name] of [[16,"slide-17-after.png"],[17,"slide-18-after.png"]]) {
  const png = await presentation.export({slide:presentation.slides.getItem(index),format:"png",scale:1});
  await fs.writeFile(`${outDir}\\${name}`,Buffer.from(await png.arrayBuffer()));
}
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
