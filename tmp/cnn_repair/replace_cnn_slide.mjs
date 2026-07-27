import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const asset="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\add_cnn_slide\\assets\\cnn_banner.png";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_repair\\edited.pptx";
const preview="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_repair\\cnn-after.png";
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const slide=p.slides.getItem(25);
const layout=JSON.parse(await (await slide.export({format:"layout"})).text());
for(const element of layout.elements){if(element.aid)p.resolve(element.aid).delete();}
slide.shapes.add({geometry:"rect",position:{left:0,top:0,width:1280,height:720},fill:"#FFFFFF",line:{style:"solid",fill:"none",width:0}});
function text(value,position,style={}){const s=slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}});s.text=value;s.text.style={fontSize:style.fontSize??24,bold:style.bold??false,color:style.color??"#0B0F19",alignment:style.alignment??"left",typeface:"Calibri"};return s;}
text("Vision",{left:64,top:34,width:420,height:30},{fontSize:16,bold:true,color:"#2563EB"});
text("CNNs reuse filters to find local patterns",{left:64,top:62,width:1080,height:62},{fontSize:40,bold:true});
const bytes=await fs.readFile(asset);
slide.images.add({blob:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),contentType:"image/png",alt:"CNN pipeline from image and shared kernels through feature maps, fully connected layers, and class probabilities",fit:"contain",position:{left:120,top:165,width:1040,height:485}});
text("26",{left:1166,top:660,width:50,height:24},{fontSize:13,color:"#667085",alignment:"right"});
slide.speakerNotes.textFrame.setText("[Sources]\nCNN overview and claims about local connectivity, feature maps, pooling, and parameter sharing: https://developersbreach.com/convolution-neural-network-deep-learning/\nCNN architecture image supplied by the user: https://i0.wp.com/developersbreach.com/wp-content/uploads/2020/08/cnn_banner.png?fit=1200%2C564&ssl=1");slide.speakerNotes.setVisible(true);
const png=await p.export({slide,format:"png",scale:1});await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(p);await pptx.save(output);
