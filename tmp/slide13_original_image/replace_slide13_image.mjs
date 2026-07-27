import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE_PPTX = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added_seq2seq_added_transformer_added_gpt_added.pptx";
const FINAL_PPTX = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_slide13_original_image.pptx";
const IMAGE_PATH = "C:\\docs\\from_linear_regression_to_gpt\\ppt_build\\assets\\annealing_loss_surface_crop.png";
const TMP_DIR = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\slide13_original_image";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(TMP_DIR, { recursive: true });
  const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE_PPTX));

  const inventory = await presentation.inspect({
    kind: "slide,textbox,shape,image,table,chart,notes,layout",
    maxChars: 500000,
  });
  await fs.writeFile(`${TMP_DIR}\\source-inventory.txt`, inventory.ndjson, "utf8");

  const slide13Record = inventory.ndjson
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line))
    .find((item) => item.kind === "slide" && item.slide === 13);
  if (!slide13Record) throw new Error("Slide 13 was not found");

  const slide = presentation.resolve(slide13Record.id);
  await writeBlob(`${TMP_DIR}\\before-slide13.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP_DIR}\\before-slide13-layout.txt`, await (await slide.export({ format: "layout" })).text(), "utf8");

  const imageRecord = inventory.ndjson
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line))
    .find((item) => {
      const b = item.bbox || [];
      return item.kind === "image" && item.slide === 13 && b.length === 4
        && Math.abs(b[0] - 112) < 2 && Math.abs(b[1] - 206) < 2
        && Math.abs(b[2] - 668) < 2 && Math.abs(b[3] - 336) < 2;
    });
  if (!imageRecord) throw new Error("Slide 13 landscape image was not found");

  const image = presentation.resolve(imageRecord.id);
  const frame = image.frame;
  const geometry = image.geometry;
  const borderRadius = image.borderRadius;
  const rotation = image.rotation;
  const flipHorizontal = image.flipHorizontal;
  const flipVertical = image.flipVertical;
  const lockAspectRatio = image.lockAspectRatio;
  const bytes = await fs.readFile(IMAGE_PATH);
  const blob = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);

  image.replace({
    blob,
    contentType: "image/png",
    alt: "Original 3D loss landscape showing gradient descent from high loss to a minimum",
    fit: "cover",
  });
  image.frame = frame;
  image.geometry = geometry;
  image.borderRadius = borderRadius;
  image.rotation = rotation;
  image.flipHorizontal = flipHorizontal;
  image.flipVertical = flipVertical;
  image.lockAspectRatio = lockAspectRatio;

  await writeBlob(`${TMP_DIR}\\after-slide13.png`, await presentation.export({ slide, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP_DIR}\\after-slide13-layout.txt`, await (await slide.export({ format: "layout" })).text(), "utf8");
  const after = await presentation.inspect({
    target: { id: slide13Record.id, beforeLines: 0, afterLines: 40 },
    kind: "slide,textbox,shape,image,notes",
    maxChars: 12000,
  });
  await fs.writeFile(`${TMP_DIR}\\after-inspect.txt`, after.ndjson, "utf8");

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL_PPTX);
  console.log(FINAL_PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
