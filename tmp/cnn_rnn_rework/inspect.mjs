import fs from 'node:fs/promises';
import { FileBlob, PresentationFile } from '../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const src='C:/workspace/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx'; const out='C:/workspace/from_linear_regression_to_gpt/tmp/cnn_rnn_rework';
const p=await PresentationFile.importPptx(await FileBlob.load(src));
for(let i=20;i<34 && i<p.slides.items.length;i++){const s=p.slides.items[i]; const b=await p.export({slide:s,format:'png',scale:1}); await fs.writeFile(`${out}/slide-${i+1}.png`,new Uint8Array(await b.arrayBuffer())); const l=JSON.parse(await (await s.export({format:'layout'})).text()); console.log(JSON.stringify({n:i+1,id:s.id,title:s.title,text:l.elements.filter(e=>e.text).map(e=>e.text).join(' | ')}));}

