import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const SOURCE = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const OUTPUT = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\neural_network_basics\\edited.pptx";
const OUT_DIR = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\neural_network_basics";
const C = { ink:"#0B0F19", muted:"#667085", line:"#CBD5E1", panel:"#F8FAFC", blue:"#2563EB", green:"#10B981", orange:"#F97316", red:"#EF4444", purple:"#7C3AED", white:"#FFFFFF" };

function addText(slide, text, position, style = {}) {
  const shape = slide.shapes.add({ geometry:"textbox", position, fill:"none", line:{style:"solid", fill:"none", width:0} });
  shape.text = text;
  shape.text.style = { fontSize:style.fontSize ?? 24, bold:style.bold ?? false, color:style.color ?? C.ink, alignment:style.alignment ?? "left" };
  return shape;
}
function box(slide, text, position, opts = {}) {
  const geometry = opts.geometry ?? "roundRect";
  const config = { geometry, position, fill:opts.fill ?? C.white, line:{style:"solid", fill:opts.line ?? C.line, width:opts.lineWidth ?? 1.5} };
  if (["rect","textbox","roundRect"].includes(geometry)) config.borderRadius = "rounded-xl";
  const shape = slide.shapes.add(config);
  if (text) { shape.text = text; shape.text.style = { fontSize:opts.fontSize ?? 22, bold:opts.bold ?? true, color:opts.color ?? C.ink, alignment:opts.alignment ?? "center" }; }
  return shape;
}
function circle(slide, text, x, y, r, opts = {}) { return box(slide, text, {left:x-r,top:y-r,width:r*2,height:r*2}, {geometry:"ellipse", fill:opts.fill ?? C.white, line:opts.line ?? C.blue, lineWidth:opts.lineWidth ?? 2, fontSize:opts.fontSize ?? 22, color:opts.color ?? C.ink}); }
function line(slide, x1, y1, x2, y2, opts = {}) { return slide.shapes.add({ geometry:"line", position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||1,height:Math.abs(y2-y1)||1}, fill:"none", line:{style:"solid",fill:opts.color ?? C.line,width:opts.width ?? 2} }); }
function arrow(slide, x, y, w, h, text = "") { return box(slide, text, {left:x,top:y,width:w,height:h}, {geometry:"rightArrow",fill:"#DBEAFE",line:C.blue,fontSize:18,color:C.blue}); }
function scaffold(slide, kicker, heading) {
  slide.background.fill = C.white;
  addText(slide, kicker, {left:64,top:34,width:420,height:30}, {fontSize:16,bold:true,color:C.blue});
  addText(slide, heading, {left:64,top:62,width:1080,height:62}, {fontSize:40,bold:true});
  box(slide, "", {left:64,top:142,width:1152,height:500}, {fill:C.panel,line:"#E2E8F0"});
  slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard neural-network concepts.");
  slide.speakerNotes.setVisible(true);
}
function nodeNetwork(slide) {
  const xs = [205, 460, 700, 1015];
  const labels = ["input x", "hidden 1", "hidden 2", "output ŷ"];
  const fills = [C.white, "#DBEAFE", "#EDE9FE", "#DCFCE7"];
  const colors = [C.muted, C.blue, C.purple, C.green];
  for (let i=0; i<3; i++) arrow(slide, xs[i]+62, 337, 120, 46, i === 0 ? "w₁, b₁" : i === 1 ? "w₂, b₂" : "w₃, b₃");
  xs.forEach((x, i) => circle(slide, i === 0 ? "x" : i === 3 ? "ŷ" : `h${i}`, x, 360, 58, {fill:fills[i],line:colors[i],fontSize:32,color:colors[i]}));
  labels.forEach((label, i) => addText(slide, label, {left:xs[i]-85,top:445,width:170,height:32}, {fontSize:22,bold:true,color:colors[i],alignment:"center"}));
  addText(slide, "width = 1  ·  hidden layers = 2", {left:325,top:205,width:630,height:42}, {fontSize:30,bold:true,color:C.blue,alignment:"center"});
  addText(slide, "h₁ = φ(w₁x + b₁)     h₂ = φ(w₂h₁ + b₂)     ŷ = w₃h₂ + b₃", {left:165,top:535,width:950,height:36}, {fontSize:24,bold:true,color:C.ink,alignment:"center"});
}
function weightsBiases(slide) {
  arrow(slide, 230, 335, 145, 52, "x");
  box(slide, "z = wx + b", {left:465,top:285,width:300,height:155}, {fill:"#DBEAFE",line:C.blue,fontSize:34,color:C.blue});
  arrow(slide, 810, 335, 145, 52, "z");
  addText(slide, "weight w", {left:455,top:475,width:150,height:34}, {fontSize:26,bold:true,color:C.orange,alignment:"center"});
  addText(slide, "bias b", {left:625,top:475,width:150,height:34}, {fontSize:26,bold:true,color:C.green,alignment:"center"});
  line(slide, 525, 510, 555, 422, {color:C.orange,width:3});
  line(slide, 692, 510, 680, 422, {color:C.green,width:3});
  addText(slide, "w scales the incoming signal", {left:185,top:555,width:390,height:34}, {fontSize:23,bold:true,color:C.orange,alignment:"center"});
  addText(slide, "b shifts the threshold", {left:705,top:555,width:340,height:34}, {fontSize:23,bold:true,color:C.green,alignment:"center"});
}
function activationRole(slide) {
  box(slide, "Linear layers only", {left:165,top:215,width:335,height:65}, {fill:C.white,line:C.line,fontSize:26});
  box(slide, "w₃(w₂(w₁x + b₁) + b₂) + b₃", {left:120,top:310,width:425,height:74}, {fill:"#FEF2F2",line:C.red,fontSize:21,color:C.red});
  addText(slide, "still one linear function", {left:145,top:415,width:375,height:34}, {fontSize:23,bold:true,color:C.red,alignment:"center"});
  arrow(slide, 570, 334, 120, 48, "φ");
  box(slide, "Activation between layers", {left:730,top:215,width:365,height:65}, {fill:"#DBEAFE",line:C.blue,fontSize:26});
  box(slide, "h = φ(wx + b)", {left:790,top:310,width:245,height:74}, {fill:"#DCFCE7",line:C.green,fontSize:28,color:C.green});
  addText(slide, "bends the mapping\nand enables nonlinear patterns", {left:740,top:410,width:345,height:62}, {fontSize:23,bold:true,color:C.green,alignment:"center"});
  addText(slide, "Without an activation function, stacking layers adds no nonlinear expressiveness.", {left:170,top:545,width:940,height:36}, {fontSize:22,bold:true,color:C.ink,alignment:"center"});
}
function curve(slide, left, top, kind, color) {
  const w=190, h=165, ox=left+25, oy=top+125;
  line(slide, ox, top+20, ox, top+145, {color:"#94A3B8",width:1}); line(slide, left+15, oy, left+175, oy, {color:"#94A3B8",width:1});
  for (let i=0;i<=38;i++) {
    const x=-3+(6*i/38); let y;
    if (kind === "ReLU") y=Math.max(0,x);
    else if (kind === "Sigmoid") y=1/(1+Math.exp(-x));
    else if (kind === "tanh") y=Math.tanh(x);
    else y=0.5*x*(1+Math.tanh(Math.sqrt(2/Math.PI)*(x+0.044715*x*x*x)));
    const px=ox+x*24; const py=oy-y*34;
    circle(slide,"",px,py,2.6,{fill:color,line:color,lineWidth:0});
  }
  addText(slide, kind, {left,top:top+168,width:w,height:30}, {fontSize:24,bold:true,color,alignment:"center"});
}
function activationGallery(slide) {
  curve(slide, 125, 230, "ReLU", C.blue);
  curve(slide, 395, 230, "Sigmoid", C.orange);
  curve(slide, 665, 230, "tanh", C.green);
  curve(slide, 935, 230, "GELU", C.purple);
  addText(slide, "ReLU: simple gating", {left:125,top:470,width:190,height:30}, {fontSize:18,color:C.muted,alignment:"center"});
  addText(slide, "Sigmoid: 0 to 1", {left:395,top:470,width:190,height:30}, {fontSize:18,color:C.muted,alignment:"center"});
  addText(slide, "tanh: −1 to 1", {left:665,top:470,width:190,height:30}, {fontSize:18,color:C.muted,alignment:"center"});
  addText(slide, "GELU: smooth gating", {left:935,top:470,width:190,height:30}, {fontSize:18,color:C.muted,alignment:"center"});
  addText(slide, "Modern Transformer blocks commonly use GELU-like activations in their feed-forward layers.", {left:155,top:545,width:970,height:36}, {fontSize:22,bold:true,color:C.ink,alignment:"center"});
}
async function renumberFooters(presentation) {
  for (let index=0; index<presentation.slides.items.length; index++) {
    const slide = presentation.slides.getItem(index);
    const layout = JSON.parse(await (await slide.export({format:"layout"})).text());
    for (const element of layout.elements) {
      const [x,y,w,h] = element.bbox ?? [];
      if (element.aid && Math.abs(x-1166)<4 && Math.abs(y-660)<8 && w<=60 && h<=30) {
        const footer = presentation.resolve(element.aid);
        footer.text = String(index+1).padStart(2,"0");
        footer.text.style = {fontSize:13,color:C.muted,alignment:"right"};
      }
    }
  }
}
async function main() {
  await fs.mkdir(OUT_DIR,{recursive:true});
  const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
  let after = presentation.slides.getItem(15);
  const slides = [
    ["Neural networks", "A tiny network: width 1, depth 2", nodeNetwork],
    ["Parameters", "Weights and biases shape each layer", weightsBiases],
    ["Activation functions", "Activation functions make depth useful", activationRole],
    ["Activation functions", "Common activation functions", activationGallery],
  ];
  for (const [kicker, heading, draw] of slides) {
    const slide = presentation.slides.insert({after}).slide;
    scaffold(slide,kicker,heading); draw(slide); after=slide;
  }
  await renumberFooters(presentation);
  for (let index=16; index<=20; index++) {
    const png = await presentation.export({slide:presentation.slides.getItem(index),format:"png",scale:1});
    await fs.writeFile(`${OUT_DIR}\\slide-${index+1}.png`,Buffer.from(await png.arrayBuffer()));
  }
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT);
}
main().catch(error => { console.error(error); process.exitCode=1; });
