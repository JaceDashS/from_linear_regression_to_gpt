import { FileBlob, PresentationFile } from '../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const p=await PresentationFile.importPptx(await FileBlob.load('C:/workspace/from_linear_regression_to_gpt/tmp/cnn_rnn_rework/cnn_rnn_reworked.pptx'));
for(let i=21;i<33 && i<p.slides.items.length;i++){const s=p.slides.items[i];const l=JSON.parse(await (await s.export({format:'layout'})).text());console.log(`${i+1}\t${l.elements.filter(e=>e.text).map(e=>e.text).join(' | ')}`)}
