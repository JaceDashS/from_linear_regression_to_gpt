import fs from "node:fs/promises";
import path from "node:path";

const root = "C:/workspace/from_linear_regression_to_gpt/tmp/ai_intro_update";
const outputSlides = [
  { outputSlide: 1, sourceSlide: 1, narrativeRole: "roadmap", reuseMode: "preserve", editTargets: ["rewrite page-range labels"] },
  { outputSlide: 2, sourceSlide: 1, narrativeRole: "AI definition and capability progression", reuseMode: "duplicate-slide", editTargets: ["replace roadmap content with four-image progression"] },
  ...Array.from({ length: 69 }, (_, index) => ({ outputSlide: index + 3, sourceSlide: index + 2, narrativeRole: "preserve existing narrative", reuseMode: "preserve", editTargets: ["rewrite footer page number"] })),
];
await fs.writeFile(path.join(root, "template-frame-map.json"), `${JSON.stringify({ outputSlides, omittedSourceSlides: [] }, null, 2)}\n`, "utf8");
await fs.writeFile(path.join(root, "template-audit.txt"), "Source deck: 70 slides, 1280×720 canvas, white panels on a gray patterned title-slide background.\nReusable pattern selected: slide 1's flat canvas, blue section label, bold title, centered white rounded panel, and lower-right page marker.\nInsertion contract: duplicate slide 1, replace its local roadmap objects with the AI capability progression, retain inherited layout/background, and insert at output slide 2.\nTypography preserved from the source deck: Times New Roman in visible content, with existing Calibri page markers.\n", "utf8");
await fs.writeFile(path.join(root, "deviation-log.txt"), "Slide 2: replaced slide 1's roadmap objects with four user-supplied images and concise capability labels so the slide explains AI before the deck enters linear regression.\nSlide 1: adjusted roadmap page ranges to account for the inserted slide.\nSlides 1-71: updated page-marker values to preserve sequential numbering.\n", "utf8");
