import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\sigma_activation\\edited.pptx";
const outDir = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\sigma_activation";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));

async function elements(slide) { return JSON.parse(await (await slide.export({format:"layout"})).text()).elements; }
function replaceByText(items, oldText, newText) {
  const item = items.find(entry => entry.text === oldText);
  if (!item) throw new Error(`Could not find: ${oldText}`);
  presentation.resolve(item.aid).text.replace(oldText,newText);
}
function note(slide, text, position) {
  const shape = slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}});
  shape.text=text; shape.text.style={fontSize:18,bold:true,color:"#2563EB",alignment:"center"};
}

const network = presentation.slides.getItem(16);
replaceByText(await elements(network), "h₁ = φ(w₁x + b₁)     h₂ = φ(w₂h₁ + b₂)     ŷ = w₃h₂ + b₃", "h₁ = σ(w₁x + b₁)     h₂ = σ(w₂h₁ + b₂)     ŷ = w₃h₂ + b₃");
note(network, "σ = activation function", {left:475,top:580,width:330,height:28});

const activation = presentation.slides.getItem(18);
const activationItems = await elements(activation);
replaceByText(activationItems, "φ", "σ");
replaceByText(activationItems, "h = φ(wx + b)", "h = σ(wx + b)");
note(activation, "σ = activation function", {left:765,top:490,width:290,height:28});

for (const [index,name] of [[16,"slide-17-after.png"],[18,"slide-19-after.png"]]) {
  const png = await presentation.export({slide:presentation.slides.getItem(index),format:"png",scale:1});
  await fs.writeFile(`${outDir}\\${name}`,Buffer.from(await png.arrayBuffer()));
}
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
