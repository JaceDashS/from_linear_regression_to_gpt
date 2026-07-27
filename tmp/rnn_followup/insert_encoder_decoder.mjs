import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_followup\\encoder_decoder_verified.pptx";
const asset = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_followup\\followup.png";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };
const rnn = p.slides.getItem(23); // current page 24
const slide = rnn.duplicate();

let layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) { const [, y] = e.bbox ?? []; if (e.aid && y >= 110 && y <= 642) p.resolve(e.aid).delete(); }
for (const image of [...slide.images.items]) image.delete();
layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) {
  const [x, y] = e.bbox ?? [];
  if (!e.aid || x < 60 || x > 85) continue;
  const shape = p.resolve(e.aid);
  if (y >= 30 && y <= 60) { shape.text = "Sequences"; shape.text.style = { typeface: "Calibri", fontSize: 16, bold: true, color: C.blue, alignment: "left" }; }
  if (y >= 60 && y <= 110) { shape.text = "Encoder–decoder maps one sequence to another"; shape.text.style = { typeface: "Calibri", fontSize: 36, bold: true, color: C.ink, alignment: "left" }; shape.text.verticalAlignment = "middle"; }
}
const bytes = await fs.readFile(asset);
slide.images.add({ blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), contentType: "image/png", alt: "LSTM-based encoder-decoder sequence-to-sequence architecture", fit: "contain", position: { left: 206, top: 148, width: 868, height: 488 } });
slide.speakerNotes.textFrame.setText("[Sources]\nImage supplied by the user: https://miro.medium.com/v2/resize:fit:1100/format:webp/1*H_heJCepouMYEqd8IwF0mQ.png\nEducational synthesis based on standard encoder-decoder sequence-to-sequence concepts.");
slide.speakerNotes.setVisible(true);

for (let i = 0; i < p.slides.items.length; i++) {
  const s = p.slides.getItem(i);
  const elems = JSON.parse(await (await s.export({ format: "layout" })).text()).elements;
  for (const e of elems) {
    const [x, y, w, h] = e.bbox ?? [];
    if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) {
      const footer = p.resolve(e.aid);
      footer.text = String(i + 1);
      footer.text.style = { typeface: "Calibri", fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}
for (const n of [24, 25, 26]) { const target = p.slides.getItem(n - 1); const png = await p.export({ slide: target, format: "png", scale: 1 }); await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_followup\\slide-${n}.png`, new Uint8Array(await png.arrayBuffer())); }
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
console.log(`slides=${p.slides.items.length}`);
