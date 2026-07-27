import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const SOURCE_PPTX = "C:\\docs\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx";
const OUTPUT_PPTX = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\dllm_mercury2_update\\updated.pptx";
const TMP_DIR = "C:\\docs\\from_linear_regression_to_gpt\\tmp\\dllm_mercury2_update";
const MERCURY_URL = "https://www.inceptionlabs.ai/blog/introducing-mercury-2";

function addText(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: style.fontSize ?? 18,
    bold: style.bold ?? false,
    color: style.color ?? "#0B0F19",
    alignment: style.alignment ?? "left",
  };
  return shape;
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  await fs.mkdir(TMP_DIR, { recursive: true });
  await fs.copyFile(SOURCE_PPTX, `${TMP_DIR}\\source-backup.pptx`);

  const presentation = await PresentationFile.importPptx(await FileBlob.load(SOURCE_PPTX));
  const inventory = await presentation.inspect({
    kind: "slide,textbox,shape,image,table,chart,notes,layout",
    maxChars: 500000,
  });
  await fs.writeFile(`${TMP_DIR}\\source-inventory.txt`, inventory.ndjson, "utf8");

  const records = inventory.ndjson.split(/\r?\n/).filter(Boolean).map((line) => JSON.parse(line));
  const slide54Record = records.find((item) => item.kind === "slide" && item.slide === 54);
  const slide55Record = records.find((item) => item.kind === "slide" && item.slide === 55);
  if (!slide54Record || !slide55Record) throw new Error("Slides 54 and 55 were not found");

  const slide54 = presentation.resolve(slide54Record.id);
  const slide55 = presentation.resolve(slide55Record.id);
  await writeBlob(`${TMP_DIR}\\before-slide54.png`, await presentation.export({ slide: slide54, format: "png", scale: 2 }));
  await writeBlob(`${TMP_DIR}\\before-slide55.png`, await presentation.export({ slide: slide55, format: "png", scale: 2 }));

  const divider = slide54.shapes.add({
    geometry: "line",
    position: { left: 920, top: 238, width: 1, height: 300 },
    fill: "none",
    line: { style: "solid", fill: "#CBD5E1", width: 1.5 },
  });
  divider.name = "mercury-example-divider";

  const label = addText(slide54, "REAL-WORLD dLLM", { left: 952, top: 248, width: 210, height: 24 }, {
    fontSize: 14,
    bold: true,
    color: "#667085",
  });
  label.name = "mercury-example-label";

  const mercury = addText(slide54, "Mercury 2 ↗", { left: 952, top: 284, width: 220, height: 36 }, {
    fontSize: 26,
    bold: true,
    color: "#7C3AED",
  });
  mercury.name = "mercury-2-link";
  mercury.text.get("Mercury 2 ↗").link = { uri: MERCURY_URL, isExternal: true };

  const description = addText(slide54, "many tokens\nrefined in parallel", { left: 952, top: 336, width: 210, height: 58 }, {
    fontSize: 18,
    color: "#0B0F19",
  });
  description.name = "mercury-example-description";

  const contrast = addText(slide54, "not one token\nat a time", { left: 952, top: 420, width: 210, height: 58 }, {
    fontSize: 18,
    bold: true,
    color: "#2563EB",
  });
  contrast.name = "mercury-example-contrast";

  slide54.speakerNotes.textFrame.setText(
    `[Sources]\n` +
    `Mercury 2 official overview: ${MERCURY_URL}\n` +
    `Mercury technical report: https://arxiv.org/abs/2506.17298\n` +
    `Large Language Diffusion Models / LLaDA: https://arxiv.org/abs/2502.09992\n` +
    `Simple and Effective Masked Diffusion Language Models: https://arxiv.org/abs/2406.07524\n\n` +
    `디퓨전 모델은 완성된 데이터에 노이즈를 더하는 정방향 과정과, 노이즈에서 원래 데이터를 복원하는 역방향 과정을 학습합니다. 이미지에서는 픽셀 노이즈를 제거하고, 텍스트 dLLM에서는 [MASK] 또는 손상된 토큰을 여러 위치에서 동시에 복원합니다.\n\n` +
    `이 슬라이드는 모든 위치가 가려진 상태에서 시작해 여러 토큰을 병렬로 예측하고, 확신이 낮은 위치를 다시 수정하면서 문장을 완성하는 과정을 보여줍니다. 자기회귀 LLM처럼 항상 왼쪽에서 오른쪽으로 한 토큰씩 만들 필요가 없다는 점이 핵심입니다.\n\n` +
    `Mercury 2는 Inception이 공개한 reasoning dLLM 사례입니다. 회사 설명에 따르면 여러 토큰을 동시에 생성하고 소수의 반복 단계에서 전체 응답을 다듬는 parallel refinement 방식을 사용합니다. 링크를 눌러 공식 소개를 확인할 수 있습니다. 속도 수치는 하드웨어와 설정에 따라 달라지므로 구조적 장점과 실제 처리량을 구분해서 설명합니다.`
  );
  slide54.speakerNotes.setVisible(true);

  slide55.speakerNotes.textFrame.setText(
    `[Sources]\n` +
    `Fast Inference from Transformers via Speculative Decoding: https://arxiv.org/abs/2211.17192\n` +
    `Google Research, Looking back at speculative decoding: https://research.google/blog/looking-back-at-speculative-decoding/\n\n` +
    `스페큘레이티브 디코딩은 작은 draft model과 큰 target model을 함께 사용합니다. 먼저 draft model이 다음 토큰 여러 개를 빠르게 제안합니다. target model은 이 후보들을 한 번의 병렬 계산으로 검사합니다.\n\n` +
    `target model이 동의한 가장 긴 앞부분은 그대로 채택합니다. 처음 불일치한 위치부터는 target model의 분포로 다시 샘플링하고 다음 라운드를 시작합니다. 올바르게 구현하면 target model 단독 생성과 같은 분포를 유지하면서 직렬 대기 횟수를 줄일 수 있습니다.\n\n` +
    `실제 속도 향상은 draft model의 비용, 후보 토큰 수, target model의 수용률에 달려 있습니다. 초안이 자주 틀리면 검증 비용만 늘 수 있으므로, 작지만 target model과 행동이 비슷한 draft model이 중요합니다.`
  );
  slide55.speakerNotes.setVisible(true);

  await writeBlob(`${TMP_DIR}\\after-slide54.png`, await presentation.export({ slide: slide54, format: "png", scale: 2 }));
  await writeBlob(`${TMP_DIR}\\after-slide55.png`, await presentation.export({ slide: slide55, format: "png", scale: 2 }));
  await fs.writeFile(`${TMP_DIR}\\after-slide54-layout.txt`, await (await slide54.export({ format: "layout" })).text(), "utf8");
  await fs.writeFile(`${TMP_DIR}\\after-slide55-layout.txt`, await (await slide55.export({ format: "layout" })).text(), "utf8");

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUTPUT_PPTX);
  console.log(OUTPUT_PPTX);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
