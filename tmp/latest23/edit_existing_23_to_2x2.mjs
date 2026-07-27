import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const src="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx",out="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\latest23\\existing23_as_2x2.pptx";
const p=await PresentationFile.importPptx(await FileBlob.load(src));const s=p.slides.getItem(22);
const C={ink:"#0B0F19",muted:"#667085",blue:"#2563EB",green:"#059669",gray:"#64748B",line:"#CBD5E1",p1:"#DBEAFE",p2:"#D1FAE5"};
function tx(t,pos,st={}){const x=s.shapes.add({geometry:"textbox",position:pos,fill:"none",line:{style:"solid",fill:"none",width:0}});x.text=t;x.text.style={typeface:st.face??"Times New Roman",fontSize:st.size??19,bold:st.bold??false,color:st.color??C.ink,alignment:st.align??"center"};x.text.verticalAlignment="middle";return x;}
function node(x,y,w,h,fill,line,text1,text2="",color=C.ink){s.shapes.add({geometry:"ellipse",position:{left:x,top:y,width:w,height:h},fill,line:{style:"solid",fill:line,width:2}});tx(text1,{left:x,top:y+10,width:w,height:text2?38:h-20},{face:"Calibri",size:22,bold:true,color,align:"center"});if(text2)tx(text2,{left:x+5,top:y+53,width:w-10,height:32},{face:"Calibri",size:16,bold:true,color:C.ink,align:"center"});}
function weight(t,x,y,color){const q=s.shapes.add({geometry:"roundRect",position:{left:x,top:y,width:118,height:34},fill:"#FFFFFF",line:{style:"solid",fill:color,width:1},borderRadius:"rounded-xl"});q.text=t;q.text.style={typeface:"Calibri",fontSize:16,bold:true,color,alignment:"center"};q.text.verticalAlignment="middle";}

let layout=JSON.parse(await (await s.export({format:"layout"})).text());
const connectors=layout.elements.filter(e=>e.scope==="slide"&&e.geometry==="straightConnector1"&&e.aid);
const normal=connectors.filter(e=>!e.verticalFlip), flipped=connectors.filter(e=>e.verticalFlip);
const keepConnectors=[normal[0],normal[1],flipped[0],normal[2]].filter(Boolean);
const keepIds=new Set(keepConnectors.map(e=>e.aid));
// Remove the user's old 3–5–5–2 nodes and labels, while preserving the background, panel, header and footer.
for(const e of layout.elements){const[x,y,w,h]=e.bbox??[];if(!e.aid||e.scope!=="slide"||keepIds.has(e.aid))continue;const keepPanel=x<100&&y>=130&&y<160&&w>1100&&h>450;const keepHeader=y<140;const keepFooter=y>640;if(!keepPanel&&!keepHeader&&!keepFooter&&y>=180)p.resolve(e.aid).delete();}
// Reuse four of the user's existing thin connectors so the direct crossed-line style is retained.
const positions=[
 {left:304,top:310,width:546,height:0},
 {left:304,top:310,width:546,height:180},
 {left:304,top:310,width:546,height:180},
 {left:304,top:490,width:546,height:0},
];
const colors=[C.blue,C.green,C.blue,C.green];
for(let i=0;i<keepConnectors.length;i++){const c=p.resolve(keepConnectors[i].aid);c.position=positions[i];c.line={style:"solid",fill:colors[i],width:2};}
// Update only the existing title text.
layout=JSON.parse(await (await s.export({format:"layout"})).text());
for(const e of layout.elements){if(e.aid&&e.text==="Neural network: basic structure"){const t=p.resolve(e.aid);t.text="A 2 × 2 neural network";t.text.style={typeface:"Times New Roman",fontSize:35.33,bold:true,color:C.ink,alignment:"left"};}}

tx("input",{left:205,top:220,width:125,height:28},{size:18,color:C.muted});
tx("perceptron",{left:820,top:220,width:170,height:28},{size:18,color:C.muted});
weight("w₁₁ = 0.2",520,273,C.blue);
weight("w₂₁ = 0.7",402,362,C.green);
weight("w₁₂ = −0.4",642,362,C.blue);
weight("w₂₂ = 0.1",520,453,C.green);
node(230,273,74,74,"#FFFFFF",C.gray,"x₁ = 2","",C.ink);
node(230,453,74,74,"#FFFFFF",C.gray,"x₂ = 1","",C.ink);
node(850,255,110,110,C.p1,C.blue,"z₁","bias b₁ = 0.1",C.blue);
node(850,435,110,110,C.p2,C.green,"z₂","bias b₂ = −0.2",C.green);
tx("2 inputs",{left:180,top:565,width:175,height:30},{size:18,color:C.muted});
tx("2 perceptrons",{left:812,top:565,width:185,height:30},{size:18,color:C.muted});
tx("Four connections become the 2 × 2 weight matrix W.",{left:410,top:603,width:470,height:26},{face:"Calibri",size:18,bold:true,color:C.ink});

// Correct the visible page marker on this physical slide.
layout=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of layout.elements){const[x,y,w,h]=e.bbox??[];if(e.aid&&x>1140&&y>640&&w<80&&h<40){const f=p.resolve(e.aid);f.text="23";f.text.style={typeface:"Calibri",fontSize:13,color:C.muted,alignment:"right"};}}
s.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard dense-layer computation.");s.speakerNotes.setVisible(true);
const b=await p.export({slide:s,format:"png",scale:1});await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\latest23\\existing23_as_2x2.png",new Uint8Array(await b.arrayBuffer()));
const f=await PresentationFile.exportPptx(p);await f.save(out);console.log(`slide=${s.id}`);
