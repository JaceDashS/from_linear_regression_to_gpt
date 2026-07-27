import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const src="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx",out="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\latest23\\slide23_2x2.pptx";
const p=await PresentationFile.importPptx(await FileBlob.load(src));const s=p.slides.getItem(22);
const C={ink:"#0B0F19",muted:"#667085",blue:"#2563EB",green:"#059669",line:"#CBD5E1",panel:"#F8FAFC",input:"#FFFFFF",p1:"#DBEAFE",p2:"#D1FAE5"};
function tx(t,pos,st={}){const x=s.shapes.add({geometry:"textbox",position:pos,fill:"none",line:{style:"solid",fill:"none",width:0}});x.text=t;x.text.style={typeface:"Calibri",fontSize:st.size??20,bold:st.bold??false,color:st.color??C.ink,alignment:st.align??"left"};x.text.verticalAlignment="middle";return x;}
function box(pos,fill,line=C.line,round=true){return s.shapes.add({geometry:round?"roundRect":"rect",position:pos,fill,line:{style:"solid",fill:line,width:1},...(round?{borderRadius:"rounded-xl"}:{})});}
function line(pos,color,flip=false){return s.shapes.add({geometry:"straightConnector1",position:pos,verticalFlip:flip,line:{style:"solid",fill:color,width:2}});}
function label(t,x,y,color){box({left:x,top:y,width:112,height:34},"#FFFFFF",color,true);tx(t,{left:x,top:y,width:112,height:34},{size:16,bold:true,color,align:"center"});}
// Preserve the user's header, but replace the old 3–5–5–2 body with a clean 2→2 network.
box({left:62,top:142,width:1156,height:501},C.panel,C.line,true);
tx("2 inputs × 2 perceptrons = four independently learned weights",{left:92,top:162,width:1090,height:30},{size:19,color:C.muted});

// Connections first so all lines remain behind nodes and labels.
// The two cross-connections use separate elbow routes so all four weights remain visible.
line({left:330,top:307,width:480,height:0},C.blue,false);                // x1 -> z1: w11
line({left:330,top:487,width:480,height:0},C.green,false);               // x2 -> z2: w22
// x1 -> z2: w21, routed below the nodes
line({left:330,top:307,width:120,height:0},C.green,false);
line({left:450,top:307,width:0,height:270},C.green,false);
line({left:450,top:577,width:360,height:0},C.green,false);
line({left:810,top:487,width:0,height:90},C.green,false);
// x2 -> z1: w12, routed above the nodes
line({left:330,top:487,width:190,height:0},C.blue,false);
line({left:520,top:217,width:0,height:270},C.blue,false);
line({left:520,top:217,width:290,height:0},C.blue,false);
line({left:810,top:217,width:0,height:90},C.blue,false);

tx("input",{left:207,top:214,width:140,height:28},{size:18,bold:true,color:C.muted,align:"center"});
tx("perceptron",{left:776,top:214,width:170,height:28},{size:18,bold:true,color:C.muted,align:"center"});

// Weight values are written directly on their corresponding lines.
label("w₁₁ = 0.2",605,270,C.blue);
label("w₂₁ = 0.7",525,560,C.green);
label("w₁₂ = −0.4",620,200,C.blue);
label("w₂₂ = 0.1",605,450,C.green);

// Input nodes.
box({left:220,top:257,width:110,height:100},C.input,C.blue,true);tx("x₁ = 2",{left:220,top:257,width:110,height:100},{size:23,bold:true,color:C.blue,align:"center"});
box({left:220,top:437,width:110,height:100},C.input,C.blue,true);tx("x₂ = 1",{left:220,top:437,width:110,height:100},{size:23,bold:true,color:C.blue,align:"center"});

// Bias is owned by the perceptron, so it is written inside each node.
box({left:810,top:247,width:150,height:120},C.p1,C.blue,true);tx("z₁ = 0.1",{left:810,top:260,width:150,height:42},{size:23,bold:true,color:C.blue,align:"center"});tx("bias b₁ = 0.1",{left:810,top:310,width:150,height:34},{size:17,bold:true,color:C.ink,align:"center"});
box({left:810,top:427,width:150,height:120},C.p2,C.green,true);tx("z₂ = 1.3",{left:810,top:440,width:150,height:42},{size:23,bold:true,color:C.green,align:"center"});tx("bias b₂ = −0.2",{left:810,top:490,width:150,height:34},{size:17,bold:true,color:C.ink,align:"center"});

tx("z₁ = 0.2x₁ − 0.4x₂ + b₁",{left:980,top:271,width:205,height:46},{size:18,bold:true,color:C.blue,align:"center"});
tx("z₂ = 0.7x₁ + 0.1x₂ + b₂",{left:980,top:451,width:205,height:46},{size:18,bold:true,color:C.green,align:"center"});
tx("The four line weights become the 2 × 2 matrix W on the next slide.",{left:140,top:603,width:1000,height:26},{size:18,bold:true,color:C.ink,align:"center"});
s.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard dense-layer computation.");s.speakerNotes.setVisible(true);

// Correct the visible page marker on this user-edited slide only.
const layout=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of layout.elements){const[x,y,w,h]=e.bbox??[];if(e.aid&&x>1140&&y>640&&w<80&&h<40){const f=p.resolve(e.aid);f.text="23";f.text.style={typeface:"Calibri",fontSize:13,color:C.muted,alignment:"right"};}}
for(const n of [22,23,24]){const b=await p.export({slide:p.slides.getItem(n-1),format:"png",scale:1});await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\latest23\\final-${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f=await PresentationFile.exportPptx(p);await f.save(out);console.log(`slides=${p.slides.items.length}`);
