import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_followup\\without_training_loop.pptx";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const found = { id: "sl/l0bepsfm" };
if (!found) throw new Error("Could not locate the Training is a loop slide.");
p.resolve(found.id).delete();

for (let i = 0; i < p.slides.items.length; i++) {
  const slide = p.slides.getItem(i);
  const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [x, y, w, h] = e.bbox ?? [];
    if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) {
      const footer = p.resolve(e.aid);
      footer.text = String(i + 1);
      footer.text.style = { typeface: "Calibri", fontSize: 13, color: "#667085", alignment: "right" };
    }
  }
}
const following = p.slides.getItem(25); // page 26 after deletion
const png = await p.export({ slide: following, format: "png", scale: 1 });
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\rnn_followup\\after-removal-slide-26.png", new Uint8Array(await png.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
console.log(`removed=${found.id}; slides=${p.slides.items.length}`);
