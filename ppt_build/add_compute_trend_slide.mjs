import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const C = {
  ink: "#0B0F19",
  muted: "#667085",
  line: "#CBD5E1",
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
    line: {
      style: opts.style ?? "solid",
      fill: opts.color ?? C.line,
      width: opts.width ?? 2,
      endArrowType: opts.endArrowType,
    },
  });
}

function circle(slide, x, y, r, color, lineColor = color) {
  return slide.shapes.add({
    geometry: "ellipse",
    position: { left: x - r, top: y - r, width: r * 2, height: r * 2 },
    fill: color,
    line: { style: "solid", fill: lineColor, width: 1.5 },
  });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function title(slide) {
  addText(slide, "Scaling trend", { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, "Frontier models scale total compute", { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
  addText(slide, "not just parameter count", { left: 822, top: 108, width: 310, height: 30 }, { fontSize: 23, bold: true, color: C.red, alignment: "right" });
}

function drawBarIcon(slide, x, y, color) {
  for (let i = 0; i < 4; i++) {
    box(slide, "", { left: x + i * 18, top: y + 44 - i * 12, width: 12, height: 18 + i * 12 }, { fill: color, line: color, lineWidth: 0 });
  }
}

function drawMoEIcon(slide, x, y) {
  circle(slide, x + 45, y + 34, 14, C.purple);
  for (let i = 0; i < 6; i++) {
    const angle = (Math.PI * 2 * i) / 6;
    const ex = x + 45 + Math.cos(angle) * 44;
    const ey = y + 34 + Math.sin(angle) * 34;
    circle(slide, ex, ey, 11, i < 2 ? C.green : "#E5E7EB", i < 2 ? C.green : "#CBD5E1");
    line(slide, x + 45, y + 34, ex, ey, { color: i < 2 ? C.green : "#D8DEE8", width: i < 2 ? 2 : 1 });
  }
}

function drawRlIcon(slide, x, y) {
  circle(slide, x + 18, y + 42, 12, C.orange);
  circle(slide, x + 55, y + 28, 12, C.orange);
  circle(slide, x + 92, y + 42, 12, C.orange);
  line(slide, x + 28, y + 38, x + 45, y + 31, { color: C.orange, width: 2, endArrowType: "triangle" });
  line(slide, x + 65, y + 31, x + 82, y + 38, { color: C.orange, width: 2, endArrowType: "triangle" });
}

function drawTimeIcon(slide, x, y) {
  circle(slide, x + 48, y + 40, 38, C.white, C.blue);
  line(slide, x + 48, y + 40, x + 48, y + 18, { color: C.blue, width: 3 });
  line(slide, x + 48, y + 40, x + 69, y + 46, { color: C.blue, width: 3 });
  line(slide, x + 88, y + 40, x + 116, y + 40, { color: C.blue, width: 2.5, endArrowType: "triangle" });
}

function drawLane(slide, cfg) {
  box(slide, "", { left: cfg.x, top: 252, width: 245, height: 296 }, { fill: C.white, line: "#E2E8F0" });
  cfg.icon(slide, cfg.x + 68, 282, cfg.color);
  addText(slide, cfg.title, { left: cfg.x + 20, top: 378, width: 205, height: 34 }, { fontSize: 25, bold: true, color: cfg.color, alignment: "center" });
  addText(slide, cfg.sub, { left: cfg.x + 22, top: 422, width: 201, height: 56 }, { fontSize: 19, color: C.ink, alignment: "center" });
  addText(slide, cfg.note, { left: cfg.x + 22, top: 494, width: 201, height: 32 }, { fontSize: 17, color: C.muted, alignment: "center" });
}

function drawSlide(slide) {
  slide.background.fill = C.white;
  title(slide);
  box(slide, "", { left: 74, top: 164, width: 1132, height: 456 }, { fill: C.panel, line: "#E2E8F0" });

  addText(slide, "more useful computation", { left: 396, top: 184, width: 488, height: 34 }, { fontSize: 27, bold: true, color: C.ink, alignment: "center" });
  line(slide, 274, 226, 1006, 226, { color: C.line, width: 4, endArrowType: "triangle" });

  const lanes = [
    { x: 112, color: C.blue, title: "pretraining", sub: "bigger models\nmore data", note: "train-time compute", icon: drawBarIcon },
    { x: 386, color: C.green, title: "MoE", sub: "many params\nfew active", note: "sparse compute", icon: drawMoEIcon },
    { x: 660, color: C.orange, title: "post-training", sub: "SFT + RL\nfeedback loops", note: "alignment compute", icon: drawRlIcon },
    { x: 934, color: C.purple, title: "test-time", sub: "think longer\nbefore answer", note: "inference compute", icon: drawTimeIcon },
  ];

  for (const lane of lanes) drawLane(slide, lane);
  addText(slide, "Trend: scale the whole compute budget", { left: 320, top: 572, width: 640, height: 34 }, { fontSize: 27, bold: true, color: C.ink, alignment: "center" });
}

async function main() {
  const p = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  while (p.slides.items.length > 46) {
    p.slides.getItem(p.slides.items.length - 1).delete();
  }
  const slide = p.slides.add();
  drawSlide(slide);
  footer(slide, 47);
  slide.speakerNotes.textFrame.setText("[Sources]\nScaling Laws for Neural Language Models: https://arxiv.org/abs/2001.08361\nTraining Compute-Optimal Large Language Models: https://arxiv.org/abs/2203.15556\nMixtral of Experts: https://arxiv.org/abs/2401.04088\nDeepSeek-V3 Technical Report: https://arxiv.org/abs/2412.19437\nOpenAI, Learning to reason with LLMs: https://openai.com/index/learning-to-reason-with-llms/\nDeepSeek-R1: https://arxiv.org/abs/2501.12948");
  slide.speakerNotes.setVisible(true);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
