import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const src="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx",out="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\weighted_nn_then_matrix.pptx";
const p=await PresentationFile.importPptx(await FileBlob.load(src));
const C={ink:"#0B0F19",muted:"#667085",blue:"#2563EB",green:"#059669",line:"#CBD5E1",row1:"#DBEAFE",row2:"#D1FAE5",panel:"#F8FAFC"};
function tx(s,t,pos,st={}){const x=s.shapes.add({geometry:"textbox",position:pos,fill:"none",line:{style:"solid",fill:"none",width:0}});x.text=t;x.text.style={typeface:"Calibri",fontSize:st.size??20,bold:st.bold??false,color:st.color??C.ink,alignment:st.align??"left"};x.text.verticalAlignment="middle";return x;}
function box(s,pos,fill,line=C.line,round=true){return s.shapes.add({geometry:round?"roundRect":"rect",position:pos,fill,line:{style:"solid",fill:line,width:1},...(round?{borderRadius:"rounded-xl"}:{})});}
function node(s,label,value,x,y,fill,line){box(s,{left:x,top:y,width:92,height:92},fill,line,true);tx(s,label,{left:x,top:y+12,width:92,height:32},{size:22,bold:true,color:line,align:"center"});tx(s,`= ${value}`,{left:x,top:y+47,width:92,height:28},{size:18,bold:true,color:C.ink,align:"center"});}
function seg(s,x,y,w,h,color,width=2){s.shapes.add({geometry:"line",position:{left:x,top:y,width:w,height:h},line:{style:"solid",fill:color,width}});}
// Duplicate the Backpropagation slide so this is inserted before the matrix slide.
const s=p.slides.getItem(21).duplicate();
box(s,{left:62,top:142,width:1156,height:501},C.panel,C.line,true);
let l=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of l.elements){const[x,y]=e.bbox??[];if(!e.aid||x<60||x>90)continue;const o=p.resolve(e.aid);if(y>=30&&y<=60){o.text="Dense layer";o.text.style={typeface:"Calibri",fontSize:16,bold:true,color:C.blue,alignment:"left"};}if(y>=60&&y<=112){o.text="Four weights connect two inputs to two hidden neurons";o.text.style={typeface:"Calibri",fontSize:35,bold:true,color:C.ink,alignment:"left"};o.text.verticalAlignment="middle";}}
tx(s,"The same four numbers will become the 2 × 2 matrix W on the next slide.",{left:92,top:160,width:1090,height:30},{size:19,color:C.muted});

// Connection routes first, behind the nodes. Blue lines enter z1; green lines enter z2.
// x1 -> z1, x2 -> z2
seg(s,270,296,410,0,C.blue,3);seg(s,270,466,410,0,C.green,3);
// x2 -> z1 via upper bus
seg(s,270,466,120,0,C.blue,2);seg(s,390,296,0,170,C.blue,2);seg(s,390,296,290,0,C.blue,2);
// x1 -> z2 via lower bus
seg(s,270,296,180,0,C.green,2);seg(s,450,296,0,170,C.green,2);seg(s,450,466,230,0,C.green,2);

tx(s,"input layer",{left:148,top:210,width:130,height:28},{size:18,bold:true,color:C.muted,align:"center"});
tx(s,"hidden layer",{left:650,top:210,width:160,height:28},{size:18,bold:true,color:C.muted,align:"center"});
node(s,"x₁","2",178,250,"#EFF6FF",C.blue);node(s,"x₂","1",178,420,"#EFF6FF",C.blue);
node(s,"z₁","0.1",680,250,C.row1,C.blue);node(s,"z₂","1.3",680,420,C.row2,C.green);

// Weight labels positioned on the route they annotate.
box(s,{left:425,top:262,width:92,height:34},"#FFFFFF",C.blue,true);tx(s,"w₁₁ = 0.2",{left:425,top:262,width:92,height:34},{size:15,bold:true,color:C.blue,align:"center"});
box(s,{left:319,top:380,width:104,height:34},"#FFFFFF",C.blue,true);tx(s,"w₁₂ = −0.4",{left:319,top:380,width:104,height:34},{size:15,bold:true,color:C.blue,align:"center"});
box(s,{left:462,top:380,width:92,height:34},"#FFFFFF",C.green,true);tx(s,"w₂₁ = 0.7",{left:462,top:380,width:92,height:34},{size:15,bold:true,color:C.green,align:"center"});
box(s,{left:526,top:432,width:92,height:34},"#FFFFFF",C.green,true);tx(s,"w₂₂ = 0.1",{left:526,top:432,width:92,height:34},{size:15,bold:true,color:C.green,align:"center"});

// Biases and scalar neuron calculations.
box(s,{left:835,top:240,width:300,height:115},"#FFFFFF",C.blue,true);tx(s,"hidden neuron 1",{left:860,top:252,width:250,height:25},{size:18,bold:true,color:C.blue});tx(s,"0.2·2 + (−0.4)·1 + 0.1 = 0.1",{left:860,top:288,width:250,height:42},{size:18,bold:true});tx(s,"bias b₁ = 0.1",{left:860,top:330,width:250,height:22},{size:16,color:C.muted});
box(s,{left:835,top:385,width:300,height:115},"#FFFFFF",C.green,true);tx(s,"hidden neuron 2",{left:860,top:397,width:250,height:25},{size:18,bold:true,color:C.green});tx(s,"0.7·2 + 0.1·1 − 0.2 = 1.3",{left:860,top:433,width:250,height:42},{size:18,bold:true});tx(s,"bias b₂ = −0.2",{left:860,top:475,width:250,height:22},{size:16,color:C.muted});
tx(s,"Each hidden neuron receives both inputs; its incoming weights form one row of W.",{left:120,top:570,width:1040,height:30},{size:19,bold:true,color:C.ink,align:"center"});
s.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard dense-layer computation.");s.speakerNotes.setVisible(true);
for(let i=0;i<p.slides.items.length;i++){const q=p.slides.getItem(i),ql=JSON.parse(await (await q.export({format:"layout"})).text());for(const e of ql.elements){const[x,y,w,h]=e.bbox??[];if(e.aid&&Math.abs(x-1166)<5&&Math.abs(y-660)<9&&w<=60&&h<=30){const f=p.resolve(e.aid);f.text=String(i+1);f.text.style={typeface:"Calibri",fontSize:13,color:C.muted,alignment:"right"};}}}
for(const n of [22,23,24,25]){const b=await p.export({slide:p.slides.getItem(n-1),format:"png",scale:1});await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\weighted-${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f=await PresentationFile.exportPptx(p);await f.save(out);console.log(`slides=${p.slides.items.length}`);
