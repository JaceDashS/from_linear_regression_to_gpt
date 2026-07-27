import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\edited.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\verified.pptx";
const p = await PresentationFile.importPptx(await FileBlob.load(source));
const slide = p.slides.getItem(20);
const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
for (const e of layout.elements) {
  const [x, y] = e.bbox ?? [];
  if (!e.aid || x < 60 || x > 80) continue;
  const shape = p.resolve(e.aid);
  if (y >= 30 && y <= 40) {
    shape.text = "Learning";
    shape.text.style = { typeface: "Georgia", fontSize: 16, bold: true, color: "#2563EB", alignment: "left" };
    shape.text.verticalAlignment = "middle";
  }
  if (y >= 58 && y <= 70) {
    shape.text = "Backpropagation uses the chain rule";
    shape.text.style = { typeface: "Georgia", fontSize: 38, bold: true, color: "#0B0F19", alignment: "left" };
    shape.text.verticalAlignment = "middle";
  }
}
const png = await p.export({ slide, format: "png", scale: 1 });
await fs.writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop_v2\\after-21-fixed.png", new Uint8Array(await png.arrayBuffer()));
const pptx = await PresentationFile.exportPptx(p);
await pptx.save(output);
