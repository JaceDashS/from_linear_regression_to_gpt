import fs from "node:fs/promises";
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
  const geometry = opts.geometry ?? "roundRect";
  const config = {
    geometry,
    position: pos,
    fill: opts.fill ?? C.white,
    line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 },
  };
  if (["rect", "textbox", "roundRect"].includes(geometry)) {
    config.borderRadius = opts.radius ?? "rounded-xl";
  }
  const s = slide.shapes.add(config);
  if (text) {
    s.text = text;
    s.text.style = {
      fontSize: opts.fontSize ?? 22,
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
    line: { style: opts.style ?? "solid", fill: opts.color ?? C.line, width: opts.width ?? 2 },
  });
}

function arrow(slide, x, y, w, h, text = "", opts = {}) {
  return box(slide, text, { left: x, top: y, width: w, height: h }, {
    geometry: "rightArrow",
    fill: opts.fill ?? "#E0F2FE",
    line: opts.line ?? C.blue,
    fontSize: opts.fontSize ?? 18,
    color: opts.color ?? C.ink,
  });
}

function frame(slide) {
  box(slide, "", { left: 64, top: 142, width: 1152, height: 500 }, { fill: C.panel, line: "#E2E8F0", radius: "rounded-lg" });
}

function title(slide, heading, kicker) {
  addText(slide, kicker, { left: 64, top: 34, width: 420, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  addText(slide, heading, { left: 64, top: 62, width: 1000, height: 62 }, { fontSize: 42, bold: true });
}

function footer(slide, n) {
  addText(slide, String(n).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

function notes(slide, lines) {
  slide.speakerNotes.textFrame.setText(["[Sources]", ...lines].join("\n"));
  slide.speakerNotes.setVisible(true);
}

function addSpecialTokenSlide(p) {
  const slide = p.slides.add();
  slide.background.fill = C.white;
  title(slide, "Special tokens can change behavior", "GPT");
  frame(slide);

  const tokens = [
    ["regular", "The"],
    ["regular", "answer"],
    ["special", "<end>"],
    ["special", "<tool_call>"],
    ["special", "<image>"],
  ];
  tokens.forEach(([kind, text], i) => {
    const x = 135 + i * 205;
    const isSpecial = kind === "special";
    box(slide, text, { left: x, top: 305, width: 145, height: 72 }, {
      fill: isSpecial ? "#FEF3C7" : C.white,
      line: isSpecial ? C.orange : C.line,
      fontSize: isSpecial ? 24 : 22,
      color: isSpecial ? C.orange : C.ink,
    });
  });
  line(slide, 150, 435, 1080, 435, { color: C.line, width: 3 });
  addText(slide, "normal text tokens", { left: 120, top: 455, width: 360, height: 30 }, { fontSize: 24, bold: true, color: C.muted, alignment: "center" });
  addText(slide, "control tokens", { left: 620, top: 455, width: 430, height: 30 }, { fontSize: 24, bold: true, color: C.orange, alignment: "center" });
  addText(slide, "Some tokens are instructions to the system, not words for the user.", { left: 190, top: 540, width: 900, height: 36 }, { fontSize: 27, bold: true, color: C.blue, alignment: "center" });
  footer(slide, p.slides.items.length);
  notes(slide, [
    "Conceptual explanation of special tokens in language-model systems.",
    "OpenAI API documentation describes tool calling and finish reasons at https://platform.openai.com/docs",
  ]);
}

function addToolTokensSlide(p) {
  const slide = p.slides.add();
  slide.background.fill = C.white;
  title(slide, "A special token can trigger an external tool", "GPT");
  frame(slide);

  box(slide, "prompt", { left: 125, top: 330, width: 145, height: 70 }, { fill: C.white, fontSize: 24 });
  arrow(slide, 295, 345, 105, 42, "");
  box(slide, "model", { left: 430, top: 300, width: 170, height: 130 }, { fill: "#E0F2FE", line: C.blue, fontSize: 30, color: C.blue });
  arrow(slide, 630, 345, 115, 42, "");
  box(slide, "<tool_call>", { left: 775, top: 320, width: 170, height: 88 }, { fill: "#FEF3C7", line: C.orange, fontSize: 25, color: C.orange });
  arrow(slide, 970, 345, 105, 42, "");
  box(slide, "tool", { left: 1100, top: 330, width: 90, height: 70 }, { fill: "#DCFCE7", line: C.green, fontSize: 24, color: C.green });

  box(slide, "<end>", { left: 500, top: 505, width: 145, height: 62 }, { fill: "#FEF2F2", line: C.red, fontSize: 26, color: C.red });
  addText(slide, "stop generation", { left: 670, top: 520, width: 240, height: 34 }, { fontSize: 26, bold: true, color: C.red });
  addText(slide, "The model may emit a control token instead of ordinary text.", { left: 205, top: 195, width: 850, height: 34 }, { fontSize: 27, bold: true, color: C.blue, alignment: "center" });
  footer(slide, p.slides.items.length);
  notes(slide, [
    "Conceptual explanation of tool-call and end-of-generation control behavior.",
    "OpenAI API documentation describes tool calling and finish reasons at https://platform.openai.com/docs",
  ]);
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  addSpecialTokenSlide(presentation);
  addToolTokensSlide(presentation);
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
