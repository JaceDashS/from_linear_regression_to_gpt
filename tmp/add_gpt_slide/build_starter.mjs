import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt_with_llm_loading_svg_curves_chain_rule_equation_nn_connections_fixed_cnn_added_rnn_gru_added_seq2seq_added_transformer_added.pptx";
const TMP = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\add_gpt_slide";
const STARTER = `${TMP}\\template-starter.pptx`;

const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE));
const newSlide = presentation.slides.getItem(12).duplicate();
newSlide.moveTo(41);

const frameMap = {
  outputSlides: presentation.slides.items.map((_, index) => {
    const outputSlide = index + 1;
    const sourceSlide = outputSlide <= 41 ? outputSlide : outputSlide === 42 ? 13 : outputSlide - 1;
    return {
      outputSlide,
      sourceSlide,
      narrativeRole: outputSlide === 42 ? "GPT decoder-only Transformer architecture overview" : "preserve source slide",
      reuseMode: "duplicate-slide",
      editTargets: outputSlide === 42
        ? [
            { sourceElementId: "slide-13-eyebrow", action: "rewrite", replacement: "GPT architecture" },
            { sourceElementId: "slide-13-title", action: "rewrite", replacement: "GPT repeats a decoder-only Transformer block" },
            { sourceElementId: "slide-13-main-image", action: "replace", replacement: "user-provided GPT architecture image" },
            { sourceElementId: "slide-13-right-card", action: "rewrite", replacement: "decoder-only architecture summary" },
            { sourceElementId: "slide-13-other-content", action: "delete", replacement: "none" },
          ]
        : outputSlide === 1
          ? [{ sourceElementId: "roadmap-range-labels", action: "rewrite", replacement: "updated 57-slide ranges" }]
          : outputSlide >= 42
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
