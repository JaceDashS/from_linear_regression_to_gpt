import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_lstm_gru\\rnn_lstm_gru_verified.pptx";
const asset = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_lstm_gru\\rnn_lstm_gru.png";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const C = { ink: "#0B0F19", muted: "#667085", blue: "#2563EB" };
const cnn = p.slides.getItem(22); // current page 23
const slide = cnn.duplicate();

// Keep the CNN slide's inherited white background, title positions, and page footer.
let layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) {
  const [, y] = e.bbox ?? [];
  if (e.aid && y >= 142 && y <= 642) p.resolve(e.aid).delete();
}
for (const image of [...slide.images.items]) image.delete();
layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) {
  const [x, y] = e.bbox ?? [];
  if (!e.aid || x < 60 || x > 85) continue;
  const shape = p.resolve(e.aid);
  if (y >= 30 && y <= 60) {
    shape.text = "Sequences";
    shape.text.style = { typeface: "Calibri", fontSize: 16, bold: true, color: C.blue, alignment: "left" };
  }
  if (y >= 60 && y <= 110) {
    shape.text = "RNNs carry state through time";
    shape.text.style = { typeface: "Calibri", fontSize: 38, bold: true, color: C.ink, alignment: "left" };
    shape.text.verticalAlignment = "middle";
  }
}
function text(value, position, style = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  s.text = value;
  s.text.style = { typeface: "Calibri", fontSize: style.fontSize ?? 18, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  if (style.verticalAlignment) s.text.verticalAlignment = style.verticalAlignment;
  return s;
}
text("LSTM and GRU add gates that decide what information to keep or forget.", { left: 74, top: 138, width: 1080, height: 30 }, { fontSize: 20, color: C.muted });
const bytes = await fs.readFile(asset);
slide.images.add({ blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), contentType: "image/png", alt: "Comparison diagram of RNN, LSTM, and GRU cells", fit: "contain", position: { left: 215, top: 190, width: 850, height: 383 } });
text("RNN: hidden state     •     LSTM: memory cell + gates     •     GRU: compact gated unit", { left: 120, top: 605, width: 1040, height: 26 }, { fontSize: 17, color: C.muted, alignment: "center" });
slide.speakerNotes.textFrame.setText("[Sources]\nImage supplied by the user: https://miro.medium.com/v2/resize:fit:1100/format:webp/1*AeKfV2uAPeMg9j8HhQZk2w.png\nEducational synthesis based on standard recurrent neural-network concepts.");
slide.speakerNotes.setVisible(true);

for (let i = 0; i < p.slides.items.length; i++) {
  const s = p.slides.getItem(i);
  const elements = JSON.parse(await (await s.export({ format: "layout" })).text()).elements;
  for (const e of elements) {
    const [x, y, w, h] = e.bbox ?? [];
    if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) {
      const footer = p.resolve(e.aid);
      footer.text = String(i + 1);
      footer.text.style = { typeface: "Calibri", fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}
for (const n of [23, 24, 25]) {
  const target = p.slides.getItem(n - 1);
  const png = await p.export({ slide: target, format: "png", scale: 1 });
  await fs.writeFile(`C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_lstm_gru\\slide-${n}.png`, new Uint8Array(await png.arrayBuffer()));
}
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
console.log(`slides=${p.slides.items.length}`);
