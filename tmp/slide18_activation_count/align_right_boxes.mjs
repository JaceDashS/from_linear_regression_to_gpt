import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\edited_aligned.pptx";
const preview="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\slide-18-aligned.png";
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const slide=p.slides.getItem(17);const layout=JSON.parse(await (await slide.export({format:"layout"})).text());
function find(text){const e=layout.elements.find(x=>x.text===text);if(!e)throw new Error(`Missing ${text}`);return p.resolve(e.aid);}
const h1=find("h₁ = σ(w₁x + b₁)");
const h2=find("h₂ = σ(w₂h₁ + b₂)");
const out=find("ŷ = w₃h₂ + b₃");
const arrows=layout.elements.filter(e=>e.geometry==="rightArrow"&&(e.bbox??[])[1]>=330&&(e.bbox??[])[1]<=340).sort((a,b)=>a.bbox[0]-b.bbox[0]);
if(arrows.length<4)throw new Error("Expected network arrows were not found");
const h1Pos={left:725,top:305,width:135,height:98};const h2Pos={left:880,top:305,width:135,height:98};const outPos={left:1035,top:305,width:135,height:98};
for(const [shape,pos] of [[h1,h1Pos],[h2,h2Pos],[out,outPos]]){shape.position=pos;shape.text.verticalAlignment="middle";}
h1.text.style={fontSize:15,bold:true,color:"#10B981",alignment:"center",typeface:"Calibri"};
h2.text.style={fontSize:15,bold:true,color:"#10B981",alignment:"center",typeface:"Calibri"};
out.text.style={fontSize:15,bold:true,color:"#10B981",alignment:"center",typeface:"Calibri"};
const between=arrows.filter(e=>e.bbox[0]>700);if(between.length!==2)throw new Error("Expected two right-side arrows");
p.resolve(between[0].aid).position={left:860,top:334,width:20,height:38};
p.resolve(between[1].aid).position={left:1015,top:334,width:20,height:38};
const png=await p.export({slide,format:"png",scale:1});await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(p);await pptx.save(output);
