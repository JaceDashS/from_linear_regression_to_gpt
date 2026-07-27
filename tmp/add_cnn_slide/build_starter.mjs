import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_cnn_slide";
const STARTER = `${TMP}\\template-starter.pptx`;

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const cnnSlide = presentation.slides.getItem(12).duplicate();
cnnSlide.moveTo(18);

const frameMap = {
  outputSlides: presentation.slides.items.map((_, index) => {
    const outputSlide = index + 1;
    const sourceSlide = outputSlide <= 18 ? outputSlide : outputSlide === 19 ? 13 : outputSlide - 1;
    return {
      outputSlide,
      sourceSlide,
      narrativeRole: outputSlide === 19 ? "CNN architecture bridge" : "preserve source slide",
      reuseMode: "duplicate-slide",
      editTargets: outputSlide === 19
        ? [
            { sourceElementId: "slide-13-eyebrow", action: "rewrite", replacement: "Vision" },
            { sourceElementId: "slide-13-title", action: "rewrite", replacement: "CNNs reuse filters to find local patterns" },
            { sourceElementId: "slide-13-main-image", action: "replace", replacement: "user-provided CNN architecture image" },
            { sourceElementId: "slide-13-supporting-diagram", action: "delete", replacement: "none" },
          ]
        : outputSlide === 1
          ? [{ sourceElementId: "roadmap-range-labels", action: "rewrite", replacement: "updated 53-slide ranges" }]
          : outputSlide >= 20
            ? [{ sourceElementId: "page-marker", action: "rewrite", replacement: String(outputSlide).padStart(2, "0") }]
            : [],
    };
  }),
  omittedSourceSlides: [],
};
await fs.writeFile(`${TMP}\\template-frame-map.json`, JSON.stringify(frameMap, null, 2), "utf8");

const starter = await PresentationFile.exportPptx(presentation);
await starter.save(STARTER);
console.log(`slides=${presentation.slides.items.length}`);
console.log(`saved=${STARTER}`);
