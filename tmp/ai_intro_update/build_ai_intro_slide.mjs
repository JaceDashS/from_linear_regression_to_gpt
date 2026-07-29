import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "file:///C:/workspace/from_linear_regression_to_gpt/ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const root = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_intro_update";
const sourcePptx = path.join(root, "source.pptx");
const outputPptx = path.join(root, "modified.pptx");
const C = {
  ink: "#0B0F19", muted: "#667085", line: "#CBD5E1", panel: "#F8FAFC",
  blue: "#2563EB", green: "#10B981", orange: "#F97316", purple: "#7C3AED", white: "#FFFFFF",
};

async function imageBytes(filename) {
  const bytes = await fs.readFile(path.join(root, filename));
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function text(slide, value, position, style = {}) {
  const shape = slide.shapes.add({ geometry: "textbox", position, fill: "none", line: { style: "solid", fill: "none", width: 0 } });
  shape.text = value;
  shape.text.style = { fontSize: style.fontSize ?? 24, bold: style.bold ?? false, color: style.color ?? C.ink, alignment: style.alignment ?? "left" };
  return shape;
}

function box(slide, position, opts = {}) {
  return slide.shapes.add({ geometry: opts.geometry ?? "roundRect", position, fill: opts.fill ?? C.white, line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.5 }, borderRadius: "rounded-xl" });
}

function connector(slide, x1, y1, x2, y2) {
  return slide.shapes.add({ geometry: "line", position: { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1) || 1, height: Math.abs(y2 - y1) || 1 }, fill: "none", line: { style: "solid", fill: C.line, width: 2.5, endArrowType: "triangle" } });
}

function addFooter(slide, page) {
  text(slide, String(page).padStart(2, "0"), { left: 1166, top: 660, width: 50, height: 24 }, { fontSize: 13, color: C.muted, alignment: "right" });
}

async function renumberFooters(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 280000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const bbox = item.bbox ?? [];
    if (bbox.length === 4 && Math.abs(bbox[0] - 1166) < 4 && Math.abs(bbox[1] - 660) < 8 && item.slide) {
      const footer = presentation.resolve(item.id);
      footer.text = String(item.slide).padStart(2, "0");
      footer.text.style = { fontSize: 13, color: C.muted, alignment: "right" };
    }
  }
}

async function shiftRoadmapPageRanges(presentation) {
  const snapshot = await presentation.inspect({ kind: "textbox", maxChars: 40000 });
  for (const row of snapshot.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    if (item.slide !== 1 || !/^Slides \d+-\d+$/.test(item.textPreview ?? "")) continue;
    const [start, end] = item.textPreview.match(/\d+/g).map(Number);
    const shape = presentation.resolve(item.id);
    shape.text = `Slides ${start + 1}-${end + 1}`;
  }
}

async function main() {
  const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePptx));
  const intro = presentation.slides.getItem(0).duplicate();
  intro.moveTo(1);
  for (const shape of [...intro.shapes.items]) shape.delete();
  for (const image of [...intro.images.items]) image.delete();
  intro.background.fill = C.white;

  text(intro, "AI landscape", { left: 64, top: 34, width: 300, height: 30 }, { fontSize: 16, bold: true, color: C.blue });
  text(intro, "AI learns patterns to act in new situations", { left: 64, top: 62, width: 1120, height: 54 }, { fontSize: 40, bold: true });
  text(intro, "From fixed rules to prediction, decision-making, and creation", { left: 64, top: 116, width: 900, height: 28 }, { fontSize: 20, color: C.muted });
  box(intro, { left: 74, top: 164, width: 1132, height: 456 }, { fill: C.panel, line: "#E2E8F0" });

  const cards = [
    { image: "level-1.jpg", type: "image/jpeg", x: 110, color: C.blue, title: "Rule-based automation", sub: "robot vacuum\nfollows set rules", alt: "Robot vacuum cleaner" },
    { image: "level-2.jpg", type: "image/jpeg", x: 380, color: C.green, title: "Prediction & decision", sub: "chess AI\nchooses a move", alt: "Chess board" },
    { image: "level-3.webp", type: "image/webp", x: 650, color: C.orange, title: "Complex strategy", sub: "game AI\nacts in changing worlds", alt: "StarCraft game strategy scene" },
    { image: "level-4.webp", type: "image/webp", x: 920, color: C.purple, title: "Generative AI", sub: "ChatGPT\ncreates language", alt: "ChatGPT logo", fit: "contain" },
  ];
  for (let index = 0; index < cards.length - 1; index += 1) connector(intro, cards[index].x + 214, 330, cards[index + 1].x - 10, 330);
  for (const card of cards) {
    box(intro, { left: card.x, top: 212, width: 210, height: 316 }, { fill: C.white, line: card.color, lineWidth: 2 });
    intro.images.add({ blob: await imageBytes(card.image), contentType: card.type, alt: card.alt, fit: card.fit ?? "cover", position: { left: card.x + 16, top: 230, width: 178, height: 134 }, geometry: "roundRect", borderRadius: "rounded-lg" });
    text(intro, card.title, { left: card.x + 12, top: 390, width: 186, height: 30 }, { fontSize: 20, bold: true, color: card.color, alignment: "center" });
    text(intro, card.sub, { left: card.x + 16, top: 434, width: 178, height: 48 }, { fontSize: 17, color: C.ink, alignment: "center" });
  }
  text(intro, "AI = learning patterns from data to predict, decide, or create", { left: 230, top: 568, width: 820, height: 32 }, { fontSize: 24, bold: true, color: C.ink, alignment: "center" });
  addFooter(intro, 2);
  intro.speakerNotes.textFrame.setText(`[Sources]\nUser-supplied image sources:\n1. https://encrypted-tbn2.gstatic.com/images?q=tbn:ANd9GcTb54J7lU3TwgQ_NiIzqn22R85C3gln6w2vp8XhqfOC-notSarD\n2. https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR4dEnPxvotPgpwHovjkZjCA84CbNgn3y91CnPf70ilYZFwjDBmSDNJBGsv&s=10\n3. https://blz-contentstack-images.akamaized.net/v3/assets/blt9c12f249ac15c7ec/blt16634e8fff6b277a/6964206ad1e07d6757e7b13d/overview_thumbnail.webp\n4. https://www.internetmatters.org/wp-content/uploads/2025/06/Chat-GPT-logo.webp\nEducational synthesis: AI learns patterns from data and applies them to prediction, decision-making, or generation.`);
  intro.speakerNotes.setVisible(true);
  await shiftRoadmapPageRanges(presentation);
  await renumberFooters(presentation);
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(outputPptx);
}

main().catch((error) => { console.error(error); process.exitCode = 1; });
