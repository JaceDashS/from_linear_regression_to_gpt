import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const source="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\edited_fully_aligned.pptx";
const preview="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\slide18_activation_count\\slide-18-fully-aligned.png";
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const slide=p.slides.getItem(17);const layout=JSON.parse(await (await slide.export({format:"layout"})).text());
function find(text){const e=layout.elements.find(x=>x.text===text);if(!e)throw new Error(`Missing ${text}`);return p.resolve(e.aid);}
const z1=find("z₁ = w₁x + b₁"),z2=find("z₂ = w₂z₁ + b₂"),out=find("ŷ = w₃z₂ + b₃");
const arrows=layout.elements.filter(e=>e.geometry==="rightArrow"&&(e.bbox??[])[1]>=330&&(e.bbox??[])[1]<=340&&(e.bbox??[])[0]<700).sort((a,b)=>a.bbox[0]-b.bbox[0]);
if(arrows.length!==2)throw new Error("Expected two left-side arrows");
for(const [shape,pos] of [[z1,{left:110,top:305,width:135,height:98}],[z2,{left:265,top:305,width:135,height:98}],[out,{left:420,top:305,width:135,height:98}]]){shape.position=pos;shape.text.style={fontSize:15,bold:true,color:shape===out?"#EF4444":"#0B0F19",alignment:"center",typeface:"Calibri"};shape.text.verticalAlignment="middle";}
p.resolve(arrows[0].aid).position={left:245,top:334,width:20,height:38};
p.resolve(arrows[1].aid).position={left:400,top:334,width:20,height:38};
const png=await p.export({slide,format:"png",scale:1});await fs.writeFile(preview,Buffer.from(await png.arrayBuffer()));
const pptx=await PresentationFile.exportPptx(p);await pptx.save(output);
