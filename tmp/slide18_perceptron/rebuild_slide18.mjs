import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_perceptron\\edited.pptx";
const preview = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_perceptron\\slide-18-after.png";
const C = {ink:"#0B0F19",muted:"#667085",line:"#CBD5E1",blue:"#2563EB",green:"#10B981",orange:"#F97316",white:"#FFFFFF"};
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const slide = presentation.slides.getItem(17);
const layout = JSON.parse(await (await slide.export({format:"layout"})).text());

for (const element of layout.elements) {
  const [x,y,w,h] = element.bbox ?? [];
  if (element.aid && y >= 142 && y <= 642) presentation.resolve(element.aid).delete();
}
slide.shapes.add({geometry:"rect",position:{left:0,top:0,width:1280,height:720},fill:"#FFFFFF",line:{style:"solid",fill:"none",width:0}});

function text(value, position, style = {}) {
  const shape = slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}});
  shape.text=value; shape.text.style={fontSize:style.fontSize ?? 24,bold:style.bold ?? false,color:style.color ?? C.ink,alignment:style.alignment ?? "left",typeface:"Calibri"}; return shape;
}
function box(value, position, opts = {}) {
  const geometry=opts.geometry ?? "roundRect";
  const config={geometry,position,fill:opts.fill ?? C.white,line:{style:"solid",fill:opts.line ?? C.line,width:opts.lineWidth ?? 1.5}};
  if (["rect","textbox","roundRect"].includes(geometry)) config.borderRadius="rounded-xl";
  const shape = slide.shapes.add(config);
  if (value) { shape.text=value; shape.text.style={fontSize:opts.fontSize ?? 24,bold:opts.bold ?? true,color:opts.color ?? C.ink,alignment:opts.alignment ?? "center",typeface:"Calibri"}; }
  return shape;
}
function arrow(x,y,w,h,label="") { return box(label,{left:x,top:y,width:w,height:h},{geometry:"rightArrow",fill:"#DBEAFE",line:C.blue,fontSize:18,color:C.blue}); }
function circle(label,x,y,r,opts={}) { return box(label,{left:x-r,top:y-r,width:r*2,height:r*2},{geometry:"ellipse",fill:opts.fill ?? C.white,line:opts.line ?? C.line,fontSize:opts.fontSize ?? 21,color:opts.color ?? C.ink}); }
function line(x1,y1,x2,y2,color=C.line,width=2) { return slide.shapes.add({geometry:"line",position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||1,height:Math.abs(y2-y1)||1},fill:"none",line:{style:"solid",fill:color,width}}); }

text("Parameters",{left:64,top:34,width:420,height:30},{fontSize:16,bold:true,color:C.blue});
text("Weights and biases shape each layer",{left:64,top:62,width:1080,height:62},{fontSize:40,bold:true});
text("18",{left:1166,top:660,width:50,height:24},{fontSize:13,color:C.muted,alignment:"right"});
text("One perceptron combines inputs",{left:160,top:180,width:960,height:38},{fontSize:28,bold:true,color:C.blue,alignment:"center"});
const ys=[285,365,445];
ys.forEach((y,i)=>{ circle(`x${i+1}`,220,y,30,{line:C.muted,color:C.muted}); line(250,y,455,365,"#CBD5E1",2); text(`w${i+1}`,{left:330,top:y-(i===1?28:12),width:56,height:24},{fontSize:17,bold:true,color:C.orange,alignment:"center"}); });
box("",{left:455,top:295,width:290,height:140},{fill:"#DBEAFE",line:C.blue});
text("weighted sum\nz = Σ wᵢxᵢ + b",{left:475,top:310,width:250,height:80},{fontSize:29,bold:true,color:C.blue,alignment:"center"});
text("b: this perceptron's bias",{left:470,top:450,width:260,height:28},{fontSize:20,bold:true,color:C.green,alignment:"center"});
arrow(770,338,110,48,"");
box("",{left:900,top:295,width:190,height:140},{fill:"#DCFCE7",line:C.green});
text("activation\nh = σ(z)",{left:910,top:310,width:170,height:80},{fontSize:27,bold:true,color:C.green,alignment:"center"});
text("σ = activation function",{left:885,top:450,width:220,height:28},{fontSize:19,bold:true,color:C.green,alignment:"center"});
text("Each input connection has a weight; the perceptron itself owns one bias.",{left:150,top:550,width:970,height:34},{fontSize:23,bold:true,color:C.ink,alignment:"center"});

const png = await presentation.export({slide,format:"png",scale:1});
await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
