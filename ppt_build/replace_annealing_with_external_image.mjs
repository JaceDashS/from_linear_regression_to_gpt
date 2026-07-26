import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const IMG = "C:\\workspace\\Attention is all you need\\ppt_build\\assets\\annealing_loss_surface_crop.png";
const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
  axis: "#94A3B8",
  panel: "#F8FAFC",
  blue: "#2563EB",
  green: "#10B981",
  orange: "#F97316",
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

function dot(slide, x, y, r, color) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x - r, top: y - r, width: r * 2, height: r * 2 },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
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
    line: {
      style: opts.style ?? "solid",
      fill: opts.color ?? C.line,
      width: opts.width ?? 2,
      endArrowType: opts.endArrowType,
    },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function drawSchedule(slide) {
  const left = 826, top = 210, w = 336, h = 324;
  box(slide, "", { left, top, width: w, height: h }, { fill: C.white, line: "#E2E8F0" });
  addText(slide, "annealing", { left: left + 30, top: top + 30, width: 190, height: 30 }, { fontSize: 25, bold: true });
  addText(slide, "temperature / learning rate", { left: left + 30, top: top + 66, width: 250, height: 24 }, { fontSize: 17, color: C.muted });

  const ax0 = left + 58, ay0 = top + 248;
  const ax1 = left + 292, ay1 = top + 112;
  line(slide, ax0, ay0, ax0, ay1, { color: C.axis, width: 2 });
  line(slide, ax0, ay0, ax1, ay0, { color: C.axis, width: 2 });
  addText(slide, "high", { left: left + 12, top: ay1 - 10, width: 50, height: 22 }, { fontSize: 15, color: C.muted, alignment: "right" });
  addText(slide, "low", { left: left + 12, top: ay0 - 12, width: 50, height: 22 }, { fontSize: 15, color: C.muted, alignment: "right" });
  addText(slide, "time", { left: left + 134, top: ay0 + 26, width: 80, height: 22 }, { fontSize: 16, color: C.muted, alignment: "center" });

  for (let i = 0; i <= 44; i++) {
    const t = i / 44;
    const x = ax0 + 28 + t * 184;
    const y = ay1 + 30 + (1 - Math.exp(-3.2 * t)) * 94;
    dot(slide, x, y, 3.1, C.blue);
  }
  dot(slide, ax0 + 28, ay1 + 30, 8.5, C.orange);
  dot(slide, ax0 + 212, ay1 + 124, 8.5, C.green);
  addText(slide, "explore", { left: left + 70, top: top + 264, width: 90, height: 26 }, { fontSize: 19, bold: true, color: C.orange, alignment: "center" });
  addText(slide, "settle", { left: left + 210, top: top + 264, width: 80, height: 26 }, { fontSize: 19, bold: true, color: C.green, alignment: "center" });
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 360000 });
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
  const old = p.slides.getItem(12);
  old.delete();
  const slide = p.slides.insert({ after: p.slides.getItem(11) }).slide;
  slide.background.fill = C.white;
  addText(slide, "Optimization", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Annealing reduces randomness over time", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });

  const imageBytes = await fs.readFile(IMG);
  slide.images.add({
    blob: imageBytes.buffer.slice(imageBytes.byteOffset, imageBytes.byteOffset + imageBytes.byteLength),
    contentType: "image/png",
    alt: "3D loss landscape with gradient descent path from high loss to minima",
    fit: "cover",
    position: { left: 112, top: 206, width: 668, height: 336 },
    geometry: "roundRect",
    borderRadius: "rounded-xl",
  });
  addText(slide, "loss landscape", { left: 132, top: 220, width: 220, height: 30 }, { fontSize: 24, bold: true, color: C.white });
  drawSchedule(slide);
  addText(slide, "early search is wide; later updates become stable", { left: 306, top: 570, width: 680, height: 34 }, { fontSize: 26, bold: true, color: C.ink, alignment: "center" });
  footer(slide, 13);
  slide.speakerNotes.textFrame.setText("[Sources]\n3D loss landscape image provided by user: https://miro.medium.com/v2/resize:fit:1400/1*_SYK3prTMdNrTW98Uic4Pg.jpeg\nEducational synthesis based on simulated annealing and learning-rate scheduling concepts.");
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
