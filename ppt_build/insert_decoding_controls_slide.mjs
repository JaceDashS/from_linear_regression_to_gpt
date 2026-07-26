import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  axis: "#94A3B8",
  panel: "#F8FAFC",
  blue: "#2563EB",
  green: "#10B981",
  orange: "#F97316",
  red: "#EF4444",
  purple: "#7C3AED",
  white: "#FFFFFF",
};

function addText(slide, text, pos, style = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: pos,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = text;
  s.text.style = {
    fontSize: style.fontSize ?? 24,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
  };
  return s;
}

function box(slide, text, pos, opts = {}) {
  const s = slide.shapes.add({
    geometry: opts.geometry ?? "roundRect",
    position: pos,
    fill: opts.fill ?? C.white,
    line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 },
    borderRadius: "rounded-xl",
  });
  if (text) {
    s.text = text;
    s.text.style = {
      fontSize: opts.fontSize ?? 24,
      bold: opts.bold ?? true,
      color: opts.color ?? C.ink,
      alignment: opts.alignment ?? "center",
    };
  }
  return s;
}

function line(slide, x1, y1, x2, y2, opts = {}) {
  return slide.shapes.add({
    geometry: "line",
    position: {
      left: Math.min(x1, x2),
      top: Math.min(y1, y2),
      width: Math.abs(x2 - x1) || 1,
      height: Math.abs(y2 - y1) || 1,
    },
    fill: "none",
    line: { style: opts.style ?? "solid", fill: opts.color ?? C.line, width: opts.width ?? 2, endArrowType: opts.endArrowType },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function bar(slide, x, y, w, h, color, label) {
  box(slide, "", { left: x, top: y, width: w, height: h }, { fill: color, line: color, lineWidth: 0 });
  addText(slide, label, { left: x - 8, top: y + h + 10, width: w + 16, height: 24 }, { fontSize: 16, color: C.muted, alignment: "center" });
}

function drawTemperature(slide) {
  box(slide, "", { left: 108, top: 204, width: 500, height: 280 }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "temperature", { left: 142, top: 232, width: 210, height: 32 }, { fontSize: 27, bold: true, color: C.orange });
  addText(slide, "low = focused", { left: 150, top: 420, width: 160, height: 26 }, { fontSize: 18, bold: true, color: C.blue, alignment: "center" });
  addText(slide, "high = varied", { left: 398, top: 420, width: 160, height: 26 }, { fontSize: 18, bold: true, color: C.orange, alignment: "center" });

  const baseY = 386;
  const low = [112, 52, 28, 15];
  const high = [74, 61, 48, 34];
  for (let i = 0; i < low.length; i++) bar(slide, 160 + i * 34, baseY - low[i], 22, low[i], C.blue, i === 0 ? "A" : "");
  for (let i = 0; i < high.length; i++) bar(slide, 408 + i * 34, baseY - high[i], 22, high[i], C.orange, i === 0 ? "A" : "");
}

function drawTopP(slide) {
  box(slide, "", { left: 672, top: 204, width: 500, height: 280 }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "top-p", { left: 706, top: 232, width: 140, height: 32 }, { fontSize: 27, bold: true, color: C.green });
  addText(slide, "keep the smallest set whose probability mass reaches p", { left: 706, top: 268, width: 400, height: 26 }, { fontSize: 17, color: C.muted });

  const probs = [88, 68, 50, 32, 20, 12];
  const labels = ["A", "B", "C", "D", "E", "F"];
  const x0 = 730;
  const y0 = 398;
  for (let i = 0; i < probs.length; i++) {
    const keep = i < 3;
    bar(slide, x0 + i * 55, y0 - probs[i], 34, probs[i], keep ? C.green : "#E5E7EB", labels[i]);
  }
  line(slide, x0 - 8, y0 - 98, x0 + 170, y0 - 98, { color: C.green, width: 2, style: "dash" });
  addText(slide, "sample only inside nucleus", { left: 760, top: 430, width: 280, height: 26 }, { fontSize: 18, bold: true, color: C.green, alignment: "center" });
}

function knob(slide, x, title, sub, color) {
  box(slide, title, { left: x, top: 520, width: 240, height: 58 }, { fill: C.white, line: color, lineWidth: 2, fontSize: 22, color });
  addText(slide, sub, { left: x, top: 586, width: 240, height: 24 }, { fontSize: 16, color: C.muted, alignment: "center" });
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 340000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (b.length === 4 && Math.abs(b[0] - 1166) < 4 && Math.abs(b[1] - 660) < 8 && item.slide) {
      const shape = presentation.resolve(item.id);
      shape.text = String(item.slide).padStart(2, "0");
      shape.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const after = p.slides.getItem(41);
  const slide = p.slides.insert({ after }).slide;
  slide.background.fill = C.white;
  addText(slide, "Decoding", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Generation has adjustable controls", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });
  drawTemperature(slide);
  drawTopP(slide);
  knob(slide, 168, "max tokens", "how long to continue", C.blue);
  knob(slide, 520, "penalties", "reduce repetition", C.purple);
  knob(slide, 872, "stop", "where to end", C.red);
  footer(slide, 43);
  slide.speakerNotes.textFrame.setText("[Sources]\nOpenAI API Reference, Responses create parameters: https://platform.openai.com/docs/api-reference/responses/create\nOpenAI API Reference, Chat create parameters: https://platform.openai.com/docs/api-reference/chat/create\nTemperature changes distribution sharpness. Top-p nucleus sampling limits sampling to a cumulative probability mass.");
  slide.speakerNotes.setVisible(true);
  await renumberFooters(p);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
