import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const deck = await PresentationFile.importPptx(await FileBlob.load("C:/docs/from_linear_regression_to_gpt/from_linear_regression_to_gpt_with_ai_landscape.pptx"));
for (const n of [1, 37]) {
  const blob = await deck.slides.getItem(n).export({ format: "png", scale: 2 });
  await fs.writeFile(`C:/docs/from_linear_regression_to_gpt/tmp/linkedin_intro_image/verify-slide-${n + 1}.png`, new Uint8Array(await blob.arrayBuffer()));
}
console.log(deck.slides.items.length);
