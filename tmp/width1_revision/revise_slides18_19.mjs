import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\width1_revision\\edited.pptx";
const outDir="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\width1_revision";
const C={ink:"#0B0F19",muted:"#667085",line:"#CBD5E1",blue:"#2563EB",green:"#10B981",orange:"#F97316",red:"#EF4444",white:"#FFFFFF"};
const p=await PresentationFile.importPptx(await FileBlob.load(source));

function addText(slide,value,position,style={}) { const s=slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}}); s.text=value; s.text.style={fontSize:style.fontSize??24,bold:style.bold??false,color:style.color??C.ink,alignment:style.alignment??"left",typeface:"Calibri"}; return s; }
function box(slide,value,position,opts={}) { const geometry=opts.geometry??"roundRect"; const config={geometry,position,fill:opts.fill??C.white,line:{style:"solid",fill:opts.line??C.line,width:opts.lineWidth??1.5}}; if (["rect","textbox","roundRect"].includes(geometry)) config.borderRadius="rounded-xl"; const s=slide.shapes.add(config); if(value){s.text=value;s.text.style={fontSize:opts.fontSize??24,bold:opts.bold??true,color:opts.color??C.ink,alignment:opts.alignment??"center",typeface:"Calibri"};}return s; }
function line(slide,x1,y1,x2,y2,color=C.line,width=2){return slide.shapes.add({geometry:"line",position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1)||1,height:Math.abs(y2-y1)||1},fill:"none",line:{style:"solid",fill:color,width}});}
function arrow(slide,x,y,w,h,label=""){return box(slide,label,{left:x,top:y,width:w,height:h},{geometry:"rightArrow",fill:"#DBEAFE",line:C.blue,fontSize:18,color:C.blue});}
async function clearBody(slide){const layout=JSON.parse(await (await slide.export({format:"layout"})).text());for(const e of layout.elements){const[,y]=e.bbox??[];if(e.aid&&y>=142&&y<=642)p.resolve(e.aid).delete();}slide.shapes.add({geometry:"rect",position:{left:0,top:0,width:1280,height:720},fill:C.white,line:{style:"solid",fill:"none",width:0}});}
function header(slide,kicker,title,n){addText(slide,kicker,{left:64,top:34,width:420,height:30},{fontSize:16,bold:true,color:C.blue});addText(slide,title,{left:64,top:62,width:1080,height:62},{fontSize:40,bold:true});addText(slide,String(n),{left:1166,top:660,width:50,height:24},{fontSize:13,color:C.muted,alignment:"right"});}

const s18=p.slides.getItem(17); await clearBody(s18); header(s18,"Parameters","Weights and biases shape each layer",18);
addText(s18,"A width-1 layer has one input and one neuron",{left:150,top:185,width:980,height:38},{fontSize:28,bold:true,color:C.blue,alignment:"center"});
box(s18,"x",{left:175,top:320,width:110,height:82},{fill:C.white,line:C.muted,fontSize:32,color:C.muted});
arrow(s18,315,338,155,46,"weight w");
box(s18,"perceptron\nz = wx + b",{left:500,top:285,width:300,height:150},{fill:"#DBEAFE",line:C.blue,fontSize:30,color:C.blue});
addText(s18,"bias b belongs inside this perceptron",{left:485,top:455,width:330,height:30},{fontSize:20,bold:true,color:C.green,alignment:"center"});
arrow(s18,835,338,110,46,"");
box(s18,"h = σ(z)",{left:975,top:320,width:135,height:82},{fill:"#DCFCE7",line:C.green,fontSize:26,color:C.green});
addText(s18,"σ = activation function",{left:935,top:455,width:215,height:28},{fontSize:19,bold:true,color:C.green,alignment:"center"});
addText(s18,"For width 1, every layer learns one weight and one bias.",{left:160,top:550,width:960,height:34},{fontSize:23,bold:true,color:C.ink,alignment:"center"});

const s19=p.slides.getItem(18); await clearBody(s19); header(s19,"Activation functions","Why activation matters in a width-1 network",19);
addText(s19,"Same width and depth — different expressive power",{left:180,top:175,width:920,height:34},{fontSize:27,bold:true,color:C.blue,alignment:"center"});
box(s19,"No activation",{left:135,top:230,width:385,height:54},{fill:"#FEF2F2",line:C.red,fontSize:25,color:C.red});
box(s19,"With activation σ",{left:760,top:230,width:385,height:54},{fill:"#DCFCE7",line:C.green,fontSize:25,color:C.green});
const leftXs=[185,285,385,485], rightXs=[810,910,1010,1110];
leftXs.forEach((x,i)=>{box(s19,["x","h₁","h₂","ŷ"][i],{left:x-31,top:350,width:62,height:62},{fill:C.white,line:i===3?C.red:C.line,fontSize:22,color:i===3?C.red:C.ink});if(i<3)arrow(s19,x+33,360,48,40,"");});
rightXs.forEach((x,i)=>{box(s19,["x","h₁","h₂","ŷ"][i],{left:x-31,top:350,width:62,height:62},{fill:i===0?C.white:"#DCFCE7",line:i===3?C.green:C.green,fontSize:22,color:C.green});if(i<3)arrow(s19,x+33,360,48,40,"");});
addText(s19,"ŷ = ax + c\nall layers collapse to one linear mapping",{left:130,top:445,width:395,height:55},{fontSize:21,bold:true,color:C.red,alignment:"center"});
addText(s19,"h₁ = σ(w₁x + b₁)\nh₂ = σ(w₂h₁ + b₂)",{left:755,top:445,width:395,height:55},{fontSize:21,bold:true,color:C.green,alignment:"center"});
addText(s19,"σ between hidden layers is what lets a width-1 network model nonlinear behavior.",{left:155,top:555,width:970,height:34},{fontSize:22,bold:true,color:C.ink,alignment:"center"});

for(const[index,name]of[[17,"slide-18-after.png"],[18,"slide-19-after.png"]]){const png=await p.export({slide:p.slides.getItem(index),format:"png",scale:1});await fs.writeFile(`${outDir}\\${name}`,Buffer.from(await png.arrayBuffer()));}
const pptx=await PresentationFile.exportPptx(p);await pptx.save(output);
