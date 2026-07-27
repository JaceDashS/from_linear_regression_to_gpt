import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/slide15-solid-normalized.pptx";
const out = "C:/docs/from_linear_regression_to_gpt/tmp/ffn_image_update/slide15-solid-verified";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
if (presentation.slides.items.length !== 58) throw new Error("Slide count changed");
const slide = presentation.slides.getItem(14);
const solid = slide.shapes.items.find((shape) => shape.name === "solid-loss-curve");
if (!solid) throw new Error("Solid curve not found");
await fs.mkdir(out, { recursive: true });
const png = await slide.export({ format: "png", scale: 1 });
await fs.writeFile(`${out}/slide-15.png`, new Uint8Array(await png.arrayBuffer()));
await fs.writeFile(`${out}/slide-15.layout.json`, await (await slide.export({ format: "layout" })).text(), "utf8");
console.log(JSON.stringify({ slides: presentation.slides.items.length, solidCurve: true }));
