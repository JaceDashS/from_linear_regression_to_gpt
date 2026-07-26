import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const PPTX = "C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx";
const ASSET_DIR = "C:\\workspace\\Attention is all you need\\ppt_build\\assets";
const LANDSCAPE_SVG = `${ASSET_DIR}\\slide13_loss_landscape_svg.svg`;
const RIGHT_CURVE_SVG = `${ASSET_DIR}\\slide13_annealing_curve.svg`;

function landscapeSvg() {
  return `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="668" height="336" viewBox="0 0 668 336">
  <defs>
    <linearGradient id="base" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#3F3C86"/>
      <stop offset="0.38" stop-color="#285F46"/>
      <stop offset="0.7" stop-color="#9B4A94"/>
      <stop offset="1" stop-color="#263B62"/>
    </linearGradient>
    <radialGradient id="hillA" cx="42%" cy="12%" r="46%">
      <stop offset="0" stop-color="#E8A7DD" stop-opacity="0.92"/>
      <stop offset="0.55" stop-color="#7D4CA6" stop-opacity="0.36"/>
      <stop offset="1" stop-color="#7D4CA6" stop-opacity="0"/>
    </radialGradient>
    <radialGradient id="hillB" cx="18%" cy="18%" r="32%">
      <stop offset="0" stop-color="#53C65A" stop-opacity="0.78"/>
      <stop offset="1" stop-color="#53C65A" stop-opacity="0"/>
    </radialGradient>
    <radialGradient id="valley" cx="39%" cy="83%" r="36%">
      <stop offset="0" stop-color="#064F25" stop-opacity="0.9"/>
      <stop offset="1" stop-color="#064F25" stop-opacity="0"/>
    </radialGradient>
    <clipPath id="surfaceClip">
      <path d="M32 250 C34 202 69 150 122 135 C161 124 184 150 210 112 C231 82 217 21 264 0 L594 0 C606 28 616 57 616 92 C617 145 584 203 571 250 L546 336 L172 336 C149 320 107 318 71 315 C51 313 39 296 32 250 Z"/>
    </clipPath>
  </defs>
  <rect width="668" height="336" rx="12" fill="#FFFFFF"/>
  <g clip-path="url(#surfaceClip)">
    <rect width="668" height="336" fill="url(#base)"/>
    <rect width="668" height="336" fill="url(#hillA)"/>
    <rect width="668" height="336" fill="url(#hillB)"/>
    <rect width="668" height="336" fill="url(#valley)"/>
    <path d="M32 250 C34 202 69 150 122 135 C161 124 184 150 210 112 C231 82 217 21 264 0 L594 0 C606 28 616 57 616 92 C617 145 584 203 571 250 L546 336 L172 336 C149 320 107 318 71 315 C51 313 39 296 32 250 Z" fill="none" stroke="#FFFFFF" stroke-opacity="0.2" stroke-width="2"/>
  </g>
  <path d="M392 98 C435 116 439 164 402 210 C375 242 298 220 291 265 C285 302 335 324 369 327" fill="none" stroke="#93D64F" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round" stroke-dasharray="9 8"/>
  <text x="350" y="82" fill="#1E155E" font-family="Calibri, Arial, sans-serif" font-size="18" font-weight="700">HIGH LOSS</text>
  <text x="326" y="200" fill="#9BE64F" font-family="Calibri, Arial, sans-serif" font-size="20" font-weight="700">GRADIENT</text>
  <text x="326" y="225" fill="#9BE64F" font-family="Calibri, Arial, sans-serif" font-size="20" font-weight="700">DESCENT</text>
  <text x="326" y="302" fill="#D7E7E3" font-family="Calibri, Arial, sans-serif" font-size="18" font-weight="700">MINIMA</text>
  <text x="500" y="205" fill="#E3C5E2" font-family="Calibri, Arial, sans-serif" font-size="18" font-weight="700" text-anchor="middle">SMOOTH</text>
  <text x="500" y="230" fill="#E3C5E2" font-family="Calibri, Arial, sans-serif" font-size="18" font-weight="700" text-anchor="middle">MORPHOLOGY</text>
</svg>`;
}

async function addSvg(slide, path, position, alt) {
  const bytes = await fs.readFile(path);
  slide.images.add({
    blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    contentType: "image/svg+xml",
    alt,
    fit: "contain",
    position,
  });
}

async function main() {
  await fs.mkdir(ASSET_DIR, { recursive: true });
  await fs.writeFile(LANDSCAPE_SVG, landscapeSvg(), "utf8");

  const presentation = await PresentationFile.importPptx(await FileBlob.load(PPTX));
  const snap = await presentation.inspect({ kind: "image", maxChars: 400000 });
  let removed = 0;
  for (const row of snap.ndjson.split(/\r?\n/)) {
    if (!row.trim()) continue;
    const item = JSON.parse(row);
    const b = item.bbox || [];
    if (item.slide !== 13 || b.length !== 4) continue;
    const isLandscapeImage = Math.abs(b[0] - 112) < 2 && Math.abs(b[1] - 206) < 2 && Math.abs(b[2] - 668) < 2 && Math.abs(b[3] - 336) < 2;
    const isFullSlideOverlay = Math.abs(b[0]) < 0.5 && Math.abs(b[1]) < 0.5 && Math.abs(b[2] - 1280) < 0.5 && Math.abs(b[3] - 720) < 0.5;
    if (isLandscapeImage || isFullSlideOverlay) {
      presentation.resolve(item.id).delete();
      removed++;
    }
  }

  const slide = presentation.slides.getItem(12);
  await addSvg(slide, LANDSCAPE_SVG, { left: 112, top: 206, width: 668, height: 336 }, "SVG loss landscape with dotted gradient descent path");
  await addSvg(slide, RIGHT_CURVE_SVG, { left: 0, top: 0, width: 1280, height: 720 }, "Annealing decay curve as SVG");

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(PPTX);
  console.log(`removed ${removed} slide 13 raster/overlay images and inserted SVG visuals`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
