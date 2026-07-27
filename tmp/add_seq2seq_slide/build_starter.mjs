import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_seq2seq_slide";
const STARTER = `${TMP}\\template-starter.pptx`;

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const newSlide = presentation.slides.getItem(12).duplicate();
newSlide.moveTo(27);

const frameMap = {
  outputSlides: presentation.slides.items.map((_, index) => {
    const outputSlide = index + 1;
    const sourceSlide = outputSlide <= 27 ? outputSlide : outputSlide === 28 ? 13 : outputSlide - 1;
    return {
      outputSlide,
      sourceSlide,
      narrativeRole: outputSlide === 28 ? "Seq2Seq encoder-decoder context bottleneck" : "preserve source slide",
      reuseMode: "duplicate-slide",
      editTargets: outputSlide === 28
        ? [
            { sourceElementId: "slide-13-eyebrow", action: "rewrite", replacement: "Encoder-decoder" },
            { sourceElementId: "slide-13-title", action: "rewrite", replacement: "Seq2Seq compresses the input into one context" },
            { sourceElementId: "slide-13-main-image", action: "delete", replacement: "user-provided Seq2Seq image in inherited image zone" },
            { sourceElementId: "slide-13-supporting-diagram", action: "delete", replacement: "none" },
            { sourceElementId: "slide-13-takeaway", action: "rewrite", replacement: "One context vector must carry the whole input" },
          ]
        : outputSlide === 1
          ? [{ sourceElementId: "roadmap-range-labels", action: "rewrite", replacement: "updated 55-slide ranges" }]
          : outputSlide >= 28
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
