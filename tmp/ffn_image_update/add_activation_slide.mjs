import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt.pptx";
const CANDIDATE = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/activation-candidate.pptx";
const OUT = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/activation-render";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function findText(slide, text) {
  const shape = slide.shapes.items.find((item) => item.text?.toString().trim() === text);
  if (!shape) throw new Error(`Text not found: ${text}`);
  return shape;
}

function replacePageNumber(slide, oldNumber, newNumber) {
  const shape = slide.shapes.items.find((item) =>
    item.text?.toString().trim() === String(oldNumber) && (item.position?.top ?? 0) >= 640,
  );
  if (!shape) throw new Error(`Page number ${oldNumber} not found`);
  shape.text = String(newNumber);
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const sourceSlide = presentation.slides.items.find((slide) =>
  slide.shapes.items.some((shape) => shape.text?.toString().trim() === "Two sigmoid layers create a nonlinear regressor"),
);
if (!sourceSlide) throw new Error("Source slide not found");

const slide = sourceSlide.duplicate();
slide.moveTo(17);

const keep = new Set([
  findText(slide, "Neural networks"),
  findText(slide, "Two sigmoid layers create a nonlinear regressor"),
  findText(slide, "17"),
]);
const panel = slide.shapes.items.find((shape) => {
  const box = shape.position;
  return box && box.left === 64 && box.top === 142 && box.width === 1152 && box.height === 500;
});
if (!panel) throw new Error("Content panel not found");
keep.add(panel);
for (const shape of [...slide.shapes.items]) {
  if (!keep.has(shape)) shape.delete();
}
for (const image of [...slide.images.items]) image.delete();
for (const chart of [...slide.charts.items]) chart.delete();

findText(slide, "Neural networks").text = "Activation";
findText(slide, "Two sigmoid layers create a nonlinear regressor").text = "Activation functions add nonlinearity";
replacePageNumber(slide, 17, 18);

const reason = slide.shapes.add({
  geometry: "textbox",
  name: "activation-reason",
  position: { left: 105, top: 162, width: 1070, height: 56 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
reason.text = "Without activation, stacked linear layers collapse into one linear function.";
reason.text.style = { fontSize: 22, bold: true, color: "#EF4444", alignment: "center", verticalAlignment: "middle" };

const xs = Array.from({ length: 21 }, (_, i) => -5 + i * 0.5);
const functions = [
  {
    name: "Sigmoid",
    color: "#2563EB",
    values: xs.map((x) => 1 / (1 + Math.exp(-x))),
    yAxis: { min: 0, max: 1, majorUnit: 0.5 },
    caption: "0 to 1 · smooth, but saturates",
  },
  {
    name: "Tanh",
    color: "#7C3AED",
    values: xs.map((x) => Math.tanh(x)),
    yAxis: { min: -1, max: 1, majorUnit: 1 },
    caption: "−1 to 1 · zero-centered",
  },
  {
    name: "ReLU",
    color: "#10B981",
    values: xs.map((x) => Math.max(0, x)),
    yAxis: { min: 0, max: 5, majorUnit: 2.5 },
    caption: "max(0, x) · simple and widely used",
  },
];

for (let i = 0; i < functions.length; i += 1) {
  const fn = functions[i];
  const left = 95 + i * 375;
  slide.charts.add("scatter", {
    position: { left, top: 230, width: 340, height: 285 },
    title: fn.name,
    titlePlacement: "aboveChart",
    titleTextStyle: { fontSize: 20, fill: "#0B0F19", bold: true },
    series: [{
      name: fn.name,
      xValues: xs,
      values: fn.values,
      line: { style: "solid", fill: fn.color, width: 3 },
      marker: { symbol: "none" },
    }],
    hasLegend: false,
    scatterOptions: { style: "smooth" },
    xAxis: {
      min: -5,
      max: 5,
      majorUnit: 5,
      line: { style: "solid", fill: "#94A3B8", width: 1 },
      majorGridlines: { style: "solid", fill: "#E2E8F0", width: 1 },
      textStyle: { fontSize: 12, fill: "#667085" },
    },
    yAxis: {
      ...fn.yAxis,
      numberFormatCode: "0.0",
      line: { style: "solid", fill: "#94A3B8", width: 1 },
      majorGridlines: { style: "solid", fill: "#E2E8F0", width: 1 },
      textStyle: { fontSize: 12, fill: "#667085" },
    },
    chartFill: "#F8FAFC",
    chartLine: { style: "solid", fill: "#E2E8F0", width: 1 },
    plotAreaFill: "#F8FAFC",
    plotAreaLine: { style: "solid", fill: "#E2E8F0", width: 1 },
  });

  const caption = slide.shapes.add({
    geometry: "textbox",
    position: { left, top: 528, width: 340, height: 52 },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  caption.text = fn.caption;
  caption.text.style = { fontSize: 18, bold: true, color: fn.color, alignment: "center", verticalAlignment: "middle" };
}

const takeaway = slide.shapes.add({
  geometry: "textbox",
  position: { left: 210, top: 585, width: 860, height: 38 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
takeaway.text = "Nonlinearity lets depth bend and combine features instead of repeating a straight line.";
takeaway.text.style = { fontSize: 19, bold: true, color: "#0B0F19", alignment: "center", verticalAlignment: "middle" };

slide.speakerNotes.textFrame.setText([
  "[Sources]",
  "Educational synthesis based on standard neural-network concepts.",
  "The plotted values are computed from sigmoid, tanh, and ReLU definitions.",
].join("\n"));
slide.speakerNotes.setVisible(true);

for (let newNumber = 19; newNumber <= 60; newNumber += 1) {
  replacePageNumber(presentation.slides.getItem(newNumber - 1), newNumber - 1, newNumber);
}

await fs.mkdir(OUT, { recursive: true });
for (let i = 16; i <= 18; i += 1) {
  const affected = presentation.slides.getItem(i);
  await writeBlob(`${OUT}/slide-${i + 1}.png`, await affected.export({ format: "png", scale: 1 }));
  await fs.writeFile(`${OUT}/slide-${i + 1}.layout.json`, await (await affected.export({ format: "layout" })).text(), "utf8");
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(CANDIDATE);
console.log(JSON.stringify({ candidate: CANDIDATE, slides: presentation.slides.items.length }));
