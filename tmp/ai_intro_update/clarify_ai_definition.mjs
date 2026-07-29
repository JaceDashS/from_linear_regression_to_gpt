import { FileBlob, PresentationFile } from "file:///C:/workspace/from_linear_regression_to_gpt/ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:/workspace/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const target = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_intro_update/redefined.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const slide = presentation.slides.getItem(1);
const snapshot = await presentation.inspect({ kind: "textbox,notes", maxChars: 24000 });
const replacements = new Map([
  ["AI landscape", "What is AI?"],
  ["AI learns patterns to act in new situations", "Artificial Intelligence learns from data"],
  ["From fixed rules to prediction, decision-making, and creation", "It applies learned patterns to new situations."],
  ["AI = learning patterns from data to predict, decide, or create", "AI enables computers to predict, decide, or create."],
]);
for (const row of snapshot.ndjson.split(/\r?\n/)) {
  if (!row.trim()) continue;
  const item = JSON.parse(row);
  if (item.slide !== 2 || !replacements.has(item.text)) continue;
  presentation.resolve(item.id).text = replacements.get(item.text);
}
slide.speakerNotes.textFrame.setText(`[Sources]\nUser-supplied image sources:\n1. https://encrypted-tbn2.gstatic.com/images?q=tbn:ANd9GcTb54J7lU3TwgQ_NiIzqn22R85C3gln6w2vp8XhqfOC-notSarD\n2. https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcR4dEnPxvotPgpwHovjkZjCA84CbNgn3y91CnPf70ilYZFwjDBmSDNJBGsv&s=10\n3. https://blz-contentstack-images.akamaized.net/v3/assets/blt9c12f249ac15c7ec/blt16634e8fff6b277a/6964206ad1e07d6757e7b13d/overview_thumbnail.webp\n4. https://www.internetmatters.org/wp-content/uploads/2025/06/Chat-GPT-logo.webp\nEducational synthesis: AI enables computers to learn patterns from data and apply them to prediction, decision-making, or generation.`);
slide.speakerNotes.setVisible(true);
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(target);
