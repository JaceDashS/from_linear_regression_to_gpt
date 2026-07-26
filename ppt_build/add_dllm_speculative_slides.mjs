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

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function title(slide, eyebrow, headline) {
  addText(slide, eyebrow, { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, headline, { left: 64, top: 62, width: 1080, height: 62 }, { fontSize: 40, bold: true });
}

function token(slide, text, x, y, opts = {}) {
  return box(slide, text, { left: x, top: y, width: 86, height: 54 }, {
    fill: opts.fill ?? C.white,
    line: opts.line ?? C.line,
    lineWidth: opts.lineWidth ?? 1.5,
    fontSize: opts.fontSize ?? 20,
    color: opts.color ?? C.ink,
  });
}

function drawDllm(slide) {
  title(slide, "Alternative generation", "Diffusion language models (dLLMs) denoise tokens");
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });

  const rows = [
    { label: "noise", y: 230, vals: ["[MASK]", "[MASK]", "[MASK]", "[MASK]", "[MASK]", "[MASK]"], color: C.muted },
    { label: "step 1", y: 338, vals: ["The", "[MASK]", "learns", "[MASK]", "[MASK]", "."], color: C.orange },
    { label: "step 2", y: 446, vals: ["The", "model", "learns", "from", "[MASK]", "."], color: C.green },
    { label: "text", y: 554, vals: ["The", "model", "learns", "from", "data", "."], color: C.blue },
  ];

  for (const row of rows) {
    addText(slide, row.label, { left: 118, top: row.y + 10, width: 120, height: 30 }, { fontSize: 22, bold: true, color: row.color, alignment: "right" });
    for (let i = 0; i < row.vals.length; i++) {
      const isMask = row.vals[i] === "[MASK]";
      token(slide, row.vals[i], 275 + i * 105, row.y, {
        fill: isMask ? "#EEF2F7" : C.white,
        line: isMask ? "#CBD5E1" : row.color,
        color: isMask ? C.muted : C.ink,
        fontSize: isMask ? 16 : 20,
      });
    }
    if (row.y < 540) line(slide, 580, row.y + 66, 580, row.y + 96, { color: C.line, width: 2.5, endArrowType: "triangle" });
  }

  addText(slide, "not always left-to-right", { left: 794, top: 179, width: 300, height: 30 }, { fontSize: 24, bold: true, color: C.purple, alignment: "center" });
}

function drawSpeculative(slide) {
  title(slide, "Faster inference", "Speculative decoding drafts, then verifies");
  box(slide, "", { left: 74, top: 154, width: 1132, height: 474 }, { fill: C.panel, line: "#E2E8F0" });

  box(slide, "small draft model", { left: 122, top: 224, width: 250, height: 86 }, { fill: C.white, line: C.green, lineWidth: 2.2, fontSize: 25, color: C.green });
  addText(slide, "cheap guesses", { left: 122, top: 318, width: 250, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });

  const draft = ["The", "cat", "sat", "on"];
  for (let i = 0; i < draft.length; i++) token(slide, draft[i], 462 + i * 93, 240, { line: C.green, color: C.ink });
  addText(slide, "candidate tokens", { left: 492, top: 318, width: 310, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });

  box(slide, "large target model", { left: 885, top: 224, width: 250, height: 86 }, { fill: C.white, line: C.blue, lineWidth: 2.2, fontSize: 25, color: C.blue });
  addText(slide, "parallel check", { left: 885, top: 318, width: 250, height: 28 }, { fontSize: 18, color: C.muted, alignment: "center" });

  line(slide, 390, 267, 442, 267, { color: C.line, width: 3, endArrowType: "triangle" });
  line(slide, 840, 267, 882, 267, { color: C.line, width: 3, endArrowType: "triangle" });

  const y = 440;
  addText(slide, "accept prefix", { left: 214, top: y - 56, width: 220, height: 30 }, { fontSize: 24, bold: true, color: C.green, alignment: "center" });
  for (let i = 0; i < 3; i++) token(slide, draft[i], 220 + i * 93, y, { line: C.green, color: C.ink });
  addText(slide, "reject / resample", { left: 694, top: y - 56, width: 250, height: 30 }, { fontSize: 24, bold: true, color: C.orange, alignment: "center" });
  token(slide, draft[3], 764, y, { line: C.orange, color: C.ink });
  token(slide, "mat", 857, y, { line: C.blue, color: C.ink });
  line(slide, 838, y + 27, 852, y + 27, { color: C.orange, width: 2.5, endArrowType: "triangle" });

  addText(slide, "same target behavior, fewer serial waits", { left: 374, top: 570, width: 530, height: 34 }, { fontSize: 25, bold: true, color: C.ink, alignment: "center" });
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 260000 });
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
  while (p.slides.items.length > 44) {
    p.slides.getItem(p.slides.items.length - 1).delete();
  }
  const s1 = p.slides.add();
  s1.background.fill = C.white;
  drawDllm(s1);
  footer(s1, 45);
  s1.speakerNotes.textFrame.setText("[Sources]\nDiscrete Diffusion in Large Language and Multimodal Models: https://arxiv.org/abs/2506.13759\nLarge Language Diffusion Models / LLaDA: https://arxiv.org/abs/2502.09992\nSimple and Effective Masked Diffusion Language Models: https://arxiv.org/abs/2406.07524");
  s1.speakerNotes.setVisible(true);

  const s2 = p.slides.add();
  s2.background.fill = C.white;
  drawSpeculative(s2);
  footer(s2, 46);
  s2.speakerNotes.textFrame.setText("[Sources]\nFast Inference from Transformers via Speculative Decoding: https://arxiv.org/abs/2211.17192\nGoogle Research, Looking back at speculative decoding: https://research.google/blog/looking-back-at-speculative-decoding/");
  s2.speakerNotes.setVisible(true);

  await renumberFooters(p);
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((e) => {
  console.error(e);
  process.exitCode = 1;
});
