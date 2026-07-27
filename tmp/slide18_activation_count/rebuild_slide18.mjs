import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\edited.pptx";
const preview="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\slide-18-after.png";
const C={ink:"#0B0F19",muted:"#667085",line:"#CBD5E1",blue:"#2563EB",green:"#10B981",red:"#EF4444",white:"#FFFFFF"};
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const slide=p.slides.getItem(17);const layout=JSON.parse(await (await slide.export({format:"layout"})).text());for(const e of layout.elements){if(e.aid)p.resolve(e.aid).delete();}
slide.shapes.add({geometry:"rect",position:{left:0,top:0,width:1280,height:720},fill:C.white,line:{style:"solid",fill:"none",width:0}});
function text(value,position,style={}){const s=slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}});s.text=value;s.text.style={fontSize:style.fontSize??24,bold:style.bold??false,color:style.color??C.ink,alignment:style.alignment??"left",typeface:"Calibri"};return s;}
function box(value,position,opts={}){const s=slide.shapes.add({geometry:"roundRect",position,fill:opts.fill??C.white,line:{style:"solid",fill:opts.line??C.line,width:opts.lineWidth??1.5},borderRadius:"rounded-xl"});if(value){s.text=value;s.text.style={fontSize:opts.fontSize??20,bold:true,color:opts.color??C.ink,alignment:"center",typeface:"Calibri"};}return s;}
function arrow(x,y,w,h,color){return slide.shapes.add({geometry:"rightArrow",position:{left:x,top:y,width:w,height:h},fill:color==="red"?"#FEE2E2":"#DBEAFE",line:{style:"solid",fill:color==="red"?C.red:C.blue,width:1.5}});}
text("Activation functions",{left:64,top:34,width:420,height:30},{fontSize:16,bold:true,color:C.blue});
text("Activation functions make depth useful",{left:64,top:62,width:1080,height:62},{fontSize:40,bold:true});
box("",{left:64,top:142,width:1152,height:500},{fill:"#F8FAFC",line:"#E2E8F0"});
box("Linear layers only",{left:115,top:190,width:440,height:58},{fill:"#FEF2F2",line:C.red,fontSize:25,color:C.red});
box("Activation between layers",{left:725,top:190,width:440,height:58},{fill:"#ECFDF5",line:C.green,fontSize:25,color:C.green});
const leftX=[115,265,415],rightX=[725,875,1025];
const linear=["z₁ = w₁x + b₁","z₂ = w₂z₁ + b₂","ŷ = w₃z₂ + b₃"];
const nonlinear=["h₁ = σ(w₁x + b₁)","h₂ = σ(w₂h₁ + b₂)","ŷ = w₃h₂ + b₃"];
leftX.forEach((x,i)=>{box(linear[i],{left:x,top:315,width:125,height:78},{fill:C.white,line:i===2?C.red:C.line,fontSize:17,color:i===2?C.red:C.ink});if(i<2)arrow(x+128,334,20,38,"red");});
rightX.forEach((x,i)=>{box(nonlinear[i],{left:x,top:315,width:125,height:78},{fill:i<2?"#DCFCE7":C.white,line:C.green,fontSize:16,color:C.green});if(i<2)arrow(x+128,334,20,38);});
text("all three layers combine into one linear mapping",{left:105,top:435,width:460,height:34},{fontSize:21,bold:true,color:C.red,alignment:"center"});
text("σ keeps the first two hidden layers nonlinear",{left:715,top:435,width:460,height:34},{fontSize:21,bold:true,color:C.green,alignment:"center"});
text("σ = activation function",{left:790,top:480,width:310,height:30},{fontSize:20,bold:true,color:C.blue,alignment:"center"});
text("Without σ, extra depth remains only a longer linear function.",{left:160,top:555,width:960,height:32},{fontSize:23,bold:true,color:C.ink,alignment:"center"});
text("18",{left:1166,top:660,width:50,height:24},{fontSize:13,color:C.muted,alignment:"right"});
slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard neural-network activation-function concepts.");slide.speakerNotes.setVisible(true);
const png=await p.export({slide,format:"png",scale:1});await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(p);await pptx.save(output);
