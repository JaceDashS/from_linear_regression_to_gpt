import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation.pptx";
const FINAL = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\fix_slide18_nn";
const RENDER = `${TMP}\\final-render`;

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function centerY(element) {
  return element.bbox[1] + element.bbox[3] / 2;
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const slide = presentation.slides.getItem(17);
const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());

const lineElements = layout.elements.filter((element) => element.scope === "slide" && element.geometry === "line");
if (lineElements.length !== 50) throw new Error(`Expected 50 broken lines on slide 18, found ${lineElements.length}.`);

const ellipseElements = layout.elements
  .filter((element) => element.scope === "slide" && element.geometry === "ellipse")
  .sort((a, b) => a.bbox[0] - b.bbox[0] || centerY(a) - centerY(b));

const layerXs = [...new Set(ellipseElements.map((element) => element.bbox[0]))].sort((a, b) => a - b);
const layers = layerXs.map((x) => ellipseElements.filter((element) => element.bbox[0] === x).map((element) => presentation.resolve(element.aid)));
const expectedLayerSizes = [3, 5, 5, 2];
if (layers.length !== expectedLayerSizes.length || layers.some((layer, index) => layer.length !== expectedLayerSizes[index])) {
  throw new Error(`Unexpected neuron layer structure: ${layers.map((layer) => layer.length).join("-")}`);
}

for (const element of lineElements) presentation.resolve(element.aid).delete();

const connectors = [];
for (let layerIndex = 0; layerIndex < layers.length - 1; layerIndex += 1) {
  for (const source of layers[layerIndex]) {
    for (const target of layers[layerIndex + 1]) {
      const connector = slide.shapes.connect(source, target, {
        kind: "straight",
        fromSide: "right",
        toSide: "left",
        line: { style: "solid", fill: "rgba(203, 213, 225, 0.45)", width: 1 },
      });
      connector.bringToFront();
      connectors.push(connector);
    }
  }
}
if (connectors.length !== 50) throw new Error(`Expected 50 replacement connectors, created ${connectors.length}.`);
for (const layer of layers) for (const neuron of layer) neuron.bringToFront();

const frameMap = {
  outputSlides: presentation.slides.items.map((_, index) => ({
    outputSlide: index + 1,
    sourceSlide: index + 1,
    narrativeRole: index === 17 ? "neural-network connectivity" : "preserve source slide",
    reuseMode: "duplicate-slide",
    editTargets: index === 17
      ? lineElements.map((element) => ({ sourceElementId: element.aid, action: "replace", replacement: "anchored straight connector" }))
      : [],
  })),
  omittedSourceSlides: [],
};
await fs.writeFile(`${TMP}\\template-frame-map.json`, JSON.stringify(frameMap, null, 2), "utf8");

await fs.mkdir(`${RENDER}\\layouts`, { recursive: true });
for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const currentSlide = presentation.slides.getItem(index);
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${RENDER}\\${stem}.png`, await currentSlide.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${RENDER}\\layouts\\${stem}.layout.json`, await (await currentSlide.export({ format: "layout" })).text(), "utf8");
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(FINAL);
console.log(`removedLines=${lineElements.length}`);
console.log(`createdConnectors=${connectors.length}`);
console.log(`saved=${FINAL}`);
