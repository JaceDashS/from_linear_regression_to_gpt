import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const imagePath="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\activation_image\\activation_functions.png";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\activation_image\\edited.pptx";
const preview="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\activation_image\\slide-19-after.png";
const sourceUrl="https://www.researchgate.net/profile/Aaron-Stebner-2/publication/341310767/figure/fig7/AS:890211844255749@1589254451431/Common-activation-functions-in-artificial-neural-networks-NNs-that-introduce.ppm";
const presentation=await PresentationFile.importPptx(await FileBlob.load(source));
const slide=presentation.slides.getItem(18);
const layout=JSON.parse(await (await slide.export({format:"layout"})).text());

for(const element of layout.elements){
  const [x,y,w,h]=element.bbox??[];
  if(element.aid && y>=190 && y<=620) presentation.resolve(element.aid).delete();
}

const bytes=await fs.readFile(imagePath);
slide.images.add({
  blob:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),
  contentType:"image/png",
  alt:"Common activation functions: perceptron, sigmoid, tanh, ReLU, leaky ReLU, ELU, softmax, and softplus",
  fit:"contain",
  position:{left:268,top:154,width:744,height:475},
});
slide.speakerNotes.textFrame.setText(`[Sources]\nActivation-function figure supplied by the user: ${sourceUrl}`);
slide.speakerNotes.setVisible(true);
const png=await presentation.export({slide,format:"png",scale:1});
await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(presentation);
await pptx.save(output);
