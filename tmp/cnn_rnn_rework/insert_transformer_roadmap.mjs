import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const src="C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx",out="C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\transformer_roadmap.pptx";
const p=await PresentationFile.importPptx(await FileBlob.load(src));
const C={ink:"#0B0F19",muted:"#667085",blue:"#2563EB",green:"#059669",orange:"#D97706",line:"#CBD5E1",pale:"#EFF6FF"};
const steps=[
 {part:"Input",name:"Tokenizer",role:"Split text into tokens and map each token to an integer ID.",detail:"Text  →  tokens  →  token IDs",accent:C.blue},
 {part:"Input",name:"Embedding",role:"Turn each token ID into a learned dense vector.",detail:"token ID  →  embedding vector",accent:C.blue},
 {part:"Input",name:"Positional encoding",role:"Add position information so the model can distinguish word order.",detail:"token embedding  +  position signal",accent:C.blue},
 {part:"Encoder",name:"Self-attention",role:"Let every input token weigh the other input tokens for context.",detail:"each token  ↔  all input tokens",accent:C.green},
 {part:"Encoder",name:"Attention head",role:"One attention head learns one kind of token-to-token relationship.",detail:"one learned perspective",accent:C.green},
 {part:"Encoder",name:"Multi-head attention",role:"Several heads learn complementary relationships in parallel.",detail:"many perspectives  →  combined context",accent:C.green},
 {part:"Encoder",name:"Encoder block",role:"Repeat self-attention and a feed-forward network to refine representations.",detail:"self-attention  +  feed-forward",accent:C.green},
 {part:"Decoder",name:"Decoder block",role:"Use masked self-attention, cross-attention, and feed-forward layers to generate.",detail:"past outputs + encoder context",accent:C.orange},
 {part:"Output",name:"Language-model head",role:"Convert the final decoder state into probabilities for the next token.",detail:"linear layer  →  softmax  →  next token",accent:C.orange},
];
function tx(s,t,pos,st={}){const x=s.shapes.add({geometry:"textbox",position:pos,fill:"none",line:{style:"solid",fill:"none",width:0}});x.text=t;x.text.style={typeface:"Calibri",fontSize:st.size??20,bold:st.bold??false,color:st.color??C.ink,alignment:st.align??"left"};x.text.verticalAlignment="middle";return x;}
function box(s,pos,fill,line=C.line){return s.shapes.add({geometry:"roundRect",position:pos,fill,line:{style:"solid",fill:line,width:1},borderRadius:"rounded-xl"});}
async function reset(s,part,name){let l=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of l.elements){const[,y]=e.bbox??[];if(e.aid&&y>=110&&y<=642)p.resolve(e.aid).delete();}for(const i of [...s.images.items])i.delete();l=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of l.elements){const[x,y]=e.bbox??[];if(!e.aid||x<60||x>90)continue;const o=p.resolve(e.aid);if(y>=30&&y<=60){o.text=`Transformer roadmap · ${part}`;o.text.style={typeface:"Calibri",fontSize:16,bold:true,color:C.blue,alignment:"left"};}if(y>=60&&y<=112){o.text=name;o.text.style={typeface:"Calibri",fontSize:38,bold:true,color:C.ink,alignment:"left"};o.text.verticalAlignment="middle";}}}
// Transformer is currently slide 29 (0-index 28). Each duplication follows the preceding roadmap slide.
let anchor=p.slides.getItem(28);
for(let current=0;current<steps.length;current++){
 const d=steps[current];const s=anchor.duplicate();await reset(s,d.part,d.name);
 tx(s,d.role,{left:76,top:143,width:1080,height:34},{size:21,color:C.muted});
 box(s,{left:170,top:250,width:940,height:150},"#F8FAFC",d.accent);
 tx(s,d.detail,{left:205,top:282,width:870,height:70},{size:31,bold:true,color:d.accent,align:"center"});
 tx(s,`Step ${current+1} of ${steps.length}`,{left:540,top:432,width:200,height:26},{size:17,color:C.muted,align:"center"});
 // Flat progress route: current topic is filled; the others remain structural context.
 const start=128, gap=112;
 for(let i=0;i<steps.length;i++){const active=i===current;box(s,{left:start+i*gap,top:510,width:92,height:46},active?d.accent:"#FFFFFF",active?d.accent:C.line);tx(s,steps[i].name,{left:start+i*gap+5,top:517,width:82,height:30},{size:active?14:12,bold:active,color:active?"#FFFFFF":C.muted,align:"center"});if(i<steps.length-1)s.shapes.add({geometry:"line",position:{left:start+i*gap+92,top:533,width:20,height:0},line:{style:"solid",fill:C.line,width:1}});}
 tx(s,`${d.part} component`,{left:484,top:605,width:312,height:28},{size:18,bold:true,color:d.accent,align:"center"});
 s.speakerNotes.textFrame.setText("[Sources]\nEducational synthesis based on the standard Transformer architecture.");s.speakerNotes.setVisible(true);
 anchor=s;
}
for(let i=0;i<p.slides.items.length;i++){const s=p.slides.getItem(i),l=JSON.parse(await (await s.export({format:"layout"})).text());for(const e of l.elements){const[x,y,w,h]=e.bbox??[];if(e.aid&&Math.abs(x-1166)<5&&Math.abs(y-660)<9&&w<=60&&h<=30){const f=p.resolve(e.aid);f.text=String(i+1);f.text.style={typeface:"Calibri",fontSize:13,color:C.muted,alignment:"right"};}}}
for(const n of [29,30,31,32,33,34,35,36,37,38,39]){const b=await p.export({slide:p.slides.getItem(n-1),format:"png",scale:1});await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\cnn_rnn_rework\\roadmap-${n}.png`,new Uint8Array(await b.arrayBuffer()));}
const f=await PresentationFile.exportPptx(p);await f.save(out);console.log(`slides=${p.slides.items.length}`);
