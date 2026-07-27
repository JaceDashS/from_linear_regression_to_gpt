import { FileBlob, PresentationFile } from "../../ppt_build/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const p=await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\from_linear_regression_to_gpt\\from_linear_regression_to_gpt.pptx"));
const scan=await p.inspect({kind:"slide,textbox,shape,notes",search:"backprop",maxChars:30000});
console.log(scan.ndjson);
