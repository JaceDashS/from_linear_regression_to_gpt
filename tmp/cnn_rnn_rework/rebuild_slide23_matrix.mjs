import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const src="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx",out="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\slide23_rebuilt.pptx";
const p=await PresentationFile.importPptx(await FileBlob.load(src));
const C={ink:"#0B0F19",muted:"#667085",blue:"#2563EB",green:"#059669",red:"#DC2626",line:"#CBD5E1",row1:"#DBEAFE",row2:"#D1FAE5",panel:"#F8FAFC"};
function tx(s,t,pos,st={}){const x=s.shapes.add({geometry:"textbox",position:pos,fill:"none",line:{style:"solid",fill:"none",width:0}});x.text=t;x.text.style={typeface:"Calibri",fontSize:st.size??20,bold:st.bold??false,color:st.color??C.ink,alignment:st.align??"left"};x.text.verticalAlignment="middle";return x;}
function rect(s,pos,fill,line=C.line,round=false){return s.shapes.add({geometry:round?"roundRect":"rect",position:pos,fill,line:{style:"solid",fill:line,width:1},...(round?{borderRadius:"rounded-xl"}:{})});}
function cell(s,t,x,y,fill,w=72,h=52){rect(s,{left:x,top:y,width:w,height:h},fill);tx(s,t,{left:x,top:y,width:w,height:h},{size:22,bold:true,align:"center"});}
// Build from the preceding Backpropagation slide to preserve the local template, then remove the rejected old slide.
const s=p.slides.getItem(21).duplicate();
// Cover the inherited body completely; retain only the template header/footer.
rect(s,{left:62,top:142,width:1156,height:501},"#F8FAFC",C.line,true);
let l=JSON.parse(await (await s.export({format:"layout"})).text());
for(const e of l.elements){const[x,y]=e.bbox??[];if(!e.aid||x<60||x>90)continue;const o=p.resolve(e.aid);if(y>=30&&y<=60){o.text="Matrix multiplication";o.text.style={typeface:"Calibri",fontSize:16,bold:true,color:C.blue,alignment:"left"};}if(y>=60&&y<=112){o.text="Two hidden neurons compute two weighted sums at once";o.text.style={typeface:"Calibri",fontSize:35,bold:true,color:C.ink,alignment:"left"};o.text.verticalAlignment="middle";}}
tx(s,"A layer with 2 inputs and 2 neurons has four weights: a 2 × 2 matrix.",{left:92,top:159,width:1090,height:30},{size:19,color:C.muted});

// Matrix equation z = Wx + b. Row colors map directly to z1 and z2.
tx(s,"z",{left:95,top:266,width:45,height:110},{size:32,bold:true,align:"center"});tx(s,"=",{left:151,top:266,width:42,height:110},{size:30,bold:true,align:"center"});
tx(s,"W",{left:275,top:207,width:120,height:30},{size:19,bold:true,color:C.blue,align:"center"});
cell(s,"0.2",215,256,C.row1);cell(s,"−0.4",287,256,C.row1);cell(s,"0.7",215,308,C.row2);cell(s,"0.1",287,308,C.row2);
tx(s,"×",{left:373,top:267,width:46,height:98},{size:30,bold:true,align:"center"});
tx(s,"x",{left:444,top:207,width:72,height:30},{size:19,bold:true,color:C.blue,align:"center"});cell(s,"2",444,256,"#FFFFFF");cell(s,"1",444,308,"#FFFFFF");
tx(s,"+",{left:533,top:267,width:46,height:98},{size:30,bold:true,align:"center"});
tx(s,"b",{left:605,top:207,width:72,height:30},{size:19,bold:true,color:C.blue,align:"center"});cell(s,"0.1",605,256,"#FFFFFF");cell(s,"−0.2",605,308,"#FFFFFF");
tx(s,"=",{left:693,top:267,width:46,height:98},{size:30,bold:true,align:"center"});
tx(s,"z",{left:772,top:207,width:72,height:30},{size:19,bold:true,color:C.blue,align:"center"});cell(s,"0.1",772,256,C.row1);cell(s,"1.3",772,308,C.row2);

// Two neurons show what the two rows mean.
tx(s,"hidden neuron 1",{left:918,top:213,width:210,height:28},{size:18,bold:true,color:C.blue,align:"center"});
rect(s,{left:986,top:255,width:74,height:74},C.row1,C.blue,true);tx(s,"z₁",{left:986,top:255,width:74,height:74},{size:25,bold:true,color:C.blue,align:"center"});
tx(s,"hidden neuron 2",{left:918,top:350,width:210,height:28},{size:18,bold:true,color:C.green,align:"center"});
rect(s,{left:986,top:392,width:74,height:74},C.row2,C.green,true);tx(s,"z₂",{left:986,top:392,width:74,height:74},{size:25,bold:true,color:C.green,align:"center"});

// Expand each row into the scalar weighted sum for verification.
tx(s,"Row 1 → z₁ = 0.2·2 + (−0.4)·1 + 0.1 = 0.1",{left:128,top:430,width:710,height:40},{size:21,bold:true,color:C.blue});
tx(s,"Row 2 → z₂ = 0.7·2 + 0.1·1 − 0.2 = 1.3",{left:128,top:480,width:710,height:40},{size:21,bold:true,color:C.green});
tx(s,"One row of W contains all weights entering one hidden neuron.",{left:120,top:563,width:1000,height:30},{size:19,bold:true,color:C.ink,align:"center"});
s.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on standard dense-layer matrix multiplication.");s.speakerNotes.setVisible(true);

// Delete only the rejected previous version, which shifted to slide 24 after duplication.
for(let i=p.slides.items.length-1;i>=0;i--){const q=p.slides.getItem(i),ql=JSON.parse(await (await q.export({format:"layout"})).text()),t=ql.elements.filter(e=>e.text).map(e=>e.text).join(" ");if(t.includes("A 2 × 2 hidden layer is one matrix multiply"))q.delete();}
for(let i=0;i<p.slides.items.length;i++){const q=p.slides.getItem(i),ql=JSON.parse(await (await q.export({format:"layout"})).text());for(const e of ql.elements){const[x,y,w,h]=e.bbox??[];if(e.aid&&Math.abs(x-1166)<5&&Math.abs(y-660)<9&&w<=60&&h<=30){const f=p.resolve(e.aid);f.text=String(i+1);f.text.style={typeface:"Calibri",fontSize:13,color:C.muted,alignment:"right"};}}}
for(const n of [22,23,24]){const b=await p.export({slide:p.slides.getItem(n-1),format:"png",scale:1});await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\rebuild23-${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f=await PresentationFile.exportPptx(p);await f.save(out);console.log(`slides=${p.slides.items.length}`);
