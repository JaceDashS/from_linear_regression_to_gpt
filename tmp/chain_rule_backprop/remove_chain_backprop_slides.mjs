import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const source = "C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const output = "C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop\\removed_chain_backprop.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));

// Slides 21–23 are the full chain-rule/backpropagation section.
for (const index of [22, 21, 20]) presentation.slides.getItem(index).delete();

// Keep page markers synchronized with the shortened deck.
for (let i = 0; i < presentation.slides.items.length; i++) {
  const slide = presentation.slides.getItem(i);
  const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
  for (const e of layout.elements) {
    const [x, y, w, h] = e.bbox ?? [];
    if (e.aid && Math.abs(x - 1166) < 5 && Math.abs(y - 660) < 9 && w <= 60 && h <= 30) {
      const footer = presentation.resolve(e.aid);
      footer.text = String(i + 1);
      footer.text.style = { typeface: "Calibri", fontSize: 13, color: "#64748B", alignment: "right" };
    }
  }
}
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(output);
const following = presentation.slides.getItem(20);
const png = await presentation.export({ slide: following, format: "png", scale: 1 });
await (await import("node:fs/promises")).writeFile("C:\\workspace\\from_linear_regression_to_gpt\\tmp\\chain_rule_backprop\\after-removal-slide-21.png", new Uint8Array(await png.arrayBuffer()));
console.log(`slides=${presentation.slides.items.length}`);
