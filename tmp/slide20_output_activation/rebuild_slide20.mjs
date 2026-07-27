import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide20_output_activation\\edited.pptx";
const preview="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide20_output_activation\\slide-20-after.png";
const C={ink:"#0B0F19",muted:"#667085",line:"#CBD5E1",blue:"#2563EB",green:"#10B981",orange:"#F97316",red:"#EF4444",purple:"#7C3AED",white:"#FFFFFF"};
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const slide=p.slides.getItem(19);
const layout=JSON.parse(await (await slide.export({format:"layout"})).text());
for(const element of layout.elements){if(element.aid)p.resolve(element.aid).delete();}
slide.shapes.add({geometry:"rect",position:{left:0,top:0,width:1280,height:720},fill:C.white,line:{style:"solid",fill:"none",width:0}});
function text(value,position,style={}){const s=slide.shapes.add({geometry:"textbox",position,fill:"none",line:{style:"solid",fill:"none",width:0}});s.text=value;s.text.style={fontSize:style.fontSize??24,bold:style.bold??false,color:style.color??C.ink,alignment:style.alignment??"left",typeface:"Calibri"};return s;}
function box(value,position,opts={}){const geometry=opts.geometry??"roundRect";const config={geometry,position,fill:opts.fill??C.white,line:{style:"solid",fill:opts.line??C.line,width:opts.lineWidth??1.5}};if(["rect","textbox","roundRect"].includes(geometry))config.borderRadius="rounded-xl";const s=slide.shapes.add(config);if(value){s.text=value;s.text.style={fontSize:opts.fontSize??24,bold:opts.bold??true,color:opts.color??C.ink,alignment:opts.alignment??"center",typeface:"Calibri"};}return s;}
function arrow(x,y,w,h,label=""){return box(label,{left:x,top:y,width:w,height:h},{geometry:"rightArrow",fill:"#DBEAFE",line:C.blue,fontSize:18,color:C.blue});}
text("Output layer",{left:64,top:34,width:420,height:30},{fontSize:16,bold:true,color:C.blue});
text("Why this regression output has no activation",{left:64,top:62,width:1080,height:62},{fontSize:39,bold:true});
box("",{left:64,top:142,width:1152,height:500},{fill:"#F8FAFC",line:"#E2E8F0"});
text("Hidden layers use σ; the output layer is chosen for the task.",{left:155,top:172,width:970,height:34},{fontSize:27,bold:true,color:C.blue,alignment:"center"});
box("x",{left:185,top:250,width:100,height:58},{fill:C.white,line:C.muted,fontSize:25,color:C.muted});arrow(300,258,70,40,"");box("h₁ = σ(·)",{left:390,top:250,width:150,height:58},{fill:"#DBEAFE",line:C.blue,fontSize:22,color:C.blue});arrow(555,258,70,40,"");box("h₂ = σ(·)",{left:645,top:250,width:150,height:58},{fill:"#DCFCE7",line:C.green,fontSize:22,color:C.green});
const panels=[
  {x:130,title:"Regression",formula:"ŷ = w₃h₂ + b₃",note:"linear output\nany real value",color:C.orange,fill:"#FFF7ED"},
  {x:470,title:"Binary classification",formula:"p = sigmoid(z)",note:"0 to 1\nprobability",color:C.green,fill:"#ECFDF5"},
  {x:810,title:"Multiclass / GPT",formula:"p = softmax(z)",note:"probabilities\nsum to 1",color:C.purple,fill:"#F5F3FF"},
];
panels.forEach(panel=>{box(panel.title,{left:panel.x,top:365,width:285,height:55},{fill:panel.fill,line:panel.color,fontSize:23,color:panel.color});box(panel.formula,{left:panel.x,top:432,width:285,height:58},{fill:C.white,line:C.line,fontSize:22,color:C.ink});text(panel.note,{left:panel.x,top:505,width:285,height:50},{fontSize:20,bold:true,color:panel.color,alignment:"center"});});
text("Here we are predicting an unrestricted numeric value, so the final layer stays linear.",{left:150,top:575,width:980,height:30},{fontSize:21,bold:true,color:C.ink,alignment:"center"});
text("20",{left:1166,top:660,width:50,height:24},{fontSize:13,color:C.muted,alignment:"right"});
slide.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard neural-network output-layer conventions.");slide.speakerNotes.setVisible(true);
const png=await p.export({slide,format:"png",scale:1});await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(p);await pptx.save(output);
