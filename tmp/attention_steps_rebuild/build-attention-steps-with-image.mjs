import fs from 'node:fs/promises';
import { FileBlob, PresentationFile } from 'C:/Users/シンジョンス/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/node/artifact_tool.mjs';

const source = 'C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx';
const output = 'C:/docs/from_linear_regression_to_gpt/tmp/attention_steps_rebuild/build-attention-steps-with-image.pptx';
const preview = 'C:/docs/from_linear_regression_to_gpt/tmp/attention_steps_rebuild/build-attention-steps-with-image.png';
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const next = p.slides.getItem(47).duplicate();
next.moveTo(47);
const byName = (name) => next.shapes.items.find((shape) => shape.name === name);
const noLine = { style: 'solid', fill: 'none', width: 0 };

byName('Rectangle 935271156').text = 'Attention roadmap';
byName('Rectangle 251777344').text = 'Build attention in three steps';
const explanation = byName('Rectangle 1633830769');
explanation.text = [
  '1. Score — Q · Kᵀ asks which tokens match the current token.',
  '',
  '2. Weight — softmax turns those scores into weights that sum to 1.',
  '',
  '3. Mix — use the weights to blend the corresponding value vectors V.',
].join('\n');
explanation.position = { left: 680, top: 235, width: 430, height: 295 };
explanation.text.fontSize = 18;
explanation.text.verticalAlignment = 'middle';
next.images.items.find((image) => image.name === 'Graphic 1633830777')?.delete();
const caption = next.shapes.add({ geometry: 'textbox', position: { left: 145, top: 525, width: 465, height: 36 }, fill: 'none', line: noLine });
caption.text = 'Brighter cells: stronger attention weights';
caption.text.style = { fontSize: 17, color: '#53637F', alignment: 'center', verticalAlignment: 'middle' };

const image = await p.export({ slide: next, format: 'png', scale: 1.5 });
await fs.writeFile(preview, new Uint8Array(await image.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
