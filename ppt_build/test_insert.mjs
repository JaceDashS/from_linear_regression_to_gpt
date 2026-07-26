import { FileBlob, PresentationFile } from "@oai/artifact-tool";
const p = await PresentationFile.importPptx(await FileBlob.load("C:\\workspace\\Attention is all you need\\from_linear_regression_to_gpt.pptx"));
console.log('count', p.slides.items.length);
console.log('getItem4', typeof p.slides.getItem(4), p.slides.getItem(4)?.id);
const ret = p.slides.insert({ after: 4 });
console.log('ret', ret, typeof ret);
console.log('count2', p.slides.items.length);
console.log('item5', typeof p.slides.getItem(5), p.slides.getItem(5)?.id);
