import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const src="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx",out="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\hidden_2x2.pptx";
const p=await PresentationFile.importPptx(await FileBlob.load(src));const C={ink:"#0B0F19",muted:"#667085",blue:"#2563EB",green:"#059669",line:"#CBD5E1",pale:"#EFF6FF"};
function tx(s,t,pos,st={}){const x=s.shapes.add({geometry:"textbox",position:pos,fill:"none",line:{style:"solid",fill:"none",width:0}});x.text=t;x.text.style={typeface:"Calibri",fontSize:st.size??20,bold:st.bold??false,color:st.color??C.ink,alignment:st.align??"left"};x.text.verticalAlignment="middle";return x;}
function node(s,label,x,y,fill,line=C.blue){const q=s.shapes.add({geometry:"ellipse",position:{left:x,top:y,width:74,height:74},fill,line:{style:"solid",fill:line,width:2}});tx(s,label,{left:x,top:y,width:74,height:74},{size:24,bold:true,align:"center"});}
async function reset(s){let l=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of l.elements){const[,y]=e.bbox??[];if(e.aid&&y>=110&&y<=642)p.resolve(e.aid).delete();}for(const i of [...s.images.items])i.delete();l=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of l.elements){const[x,y]=e.bbox??[];if(!e.aid||x<60||x>90)continue;const o=p.resolve(e.aid);if(y>=30&&y<=60){o.text="Matrix multiplication";o.text.style={typeface:"Calibri",fontSize:16,bold:true,color:C.blue,alignment:"left"};}if(y>=60&&y<=112){o.text="A 2 × 2 hidden layer is one matrix multiply";o.text.style={typeface:"Calibri",fontSize:35,bold:true,color:C.ink,alignment:"left"};o.text.verticalAlignment="middle";}}}
// Chain-rule / backpropagation is slide 22 (index 21); duplication inserts the example immediately after it.
const s=p.slides.getItem(21).duplicate();await reset(s);
tx(s,"Two inputs feed two hidden neurons. Each edge is a weight in W.",{left:75,top:137,width:1080,height:30},{size:19,color:C.muted});
// Connectors are created before nodes.
const inputs=[[170,260],[170,400]], hiddens=[[405,260],[405,400]];
for(const [ix,iy] of inputs)for(const[hx,hy] of hiddens){const x1=ix+74,y1=iy+37,x2=hx,y2=hy+37;s.shapes.add({geometry:"line",position:{left:Math.min(x1,x2),top:Math.min(y1,y2),width:Math.abs(x2-x1),height:Math.abs(y2-y1)},line:{style:"solid",fill:"#93C5FD",width:2}});}
tx(s,"input x",{left:145,top:205,width:125,height:26},{size:17,bold:true,color:C.muted,align:"center"});tx(s,"hidden pre-activation z",{left:355,top:205,width:175,height:26},{size:17,bold:true,color:C.muted,align:"center"});
node(s,"x₁",170,260,"#EFF6FF");node(s,"x₂",170,400,"#EFF6FF");node(s,"z₁",405,260,"#ECFDF5",C.green);node(s,"z₂",405,400,"#ECFDF5",C.green);
tx(s,"W = 2 × 2 weights",{left:250,top:335,width:190,height:26},{size:16,bold:true,color:C.blue,align:"center"});
// Formula panel.
s.shapes.add({geometry:"roundRect",position:{left:610,top:205,width:560,height:330},fill:"#F8FAFC",line:{style:"solid",fill:C.line,width:1},borderRadius:"rounded-xl"});
tx(s,"Forward pass",{left:645,top:228,width:200,height:30},{size:22,bold:true,color:C.blue});
tx(s,"z = Wx + b",{left:784,top:270,width:230,height:42},{size:28,bold:true,color:C.ink,align:"center"});
tx(s,"[ z₁ ]     [ 0.2  −0.4 ] [ x₁ ]     [ 0.1 ]",{left:652,top:326,width:476,height:34},{size:22,align:"center"});
tx(s,"[ z₂ ]  =  [ 0.7   0.1 ] [ x₂ ]  +  [−0.2]",{left:652,top:367,width:476,height:34},{size:22,align:"center"});
tx(s,"x = [2, 1]ᵀ  →  z = [0.1, 1.3]ᵀ",{left:675,top:432,width:430,height:34},{size:20,bold:true,color:C.green,align:"center"});
tx(s,"Then apply the activation:  a = σ(z)",{left:685,top:478,width:410,height:28},{size:18,color:C.muted,align:"center"});
tx(s,"Rows of W correspond to hidden neurons; columns correspond to input features.",{left:115,top:590,width:1050,height:28},{size:18,bold:true,color:C.blue,align:"center"});
s.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard dense neural-network matrix operations.");s.speakerNotes.setVisible(true);
for(let i=0;i<p.slides.items.length;i++){const q=p.slides.getItem(i),l=JSON.parse(await (await q.export({format:"layout"})).text());for(const e of l.elements){const[x,y,w,h]=e.bbox??[];if(e.aid&&Math.abs(x-1166)<5&&Math.abs(y-660)<9&&w<=60&&h<=30){const f=p.resolve(e.aid);f.text=String(i+1);f.text.style={typeface:"Calibri",fontSize:13,color:C.muted,alignment:"right"};}}}
for(const n of [22,23,24]){const b=await p.export({slide:p.slides.getItem(n-1),format:"png",scale:1});await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\matrix-${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f=await PresentationFile.exportPptx(p);await f.save(out);console.log(`slides=${p.slides.items.length}`);
