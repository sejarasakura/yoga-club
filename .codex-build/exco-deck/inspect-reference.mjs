import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "/Users/susielim/Documents/ChatGPT/YOGA Hosting/1nd EXCO proposal meeting 2026.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const snapshot = await presentation.inspect({
  kind: "deck,slide,textbox,shape,image,table,chart,notes,layout",
  include: "id,slide,name,title,text,textPreview,bbox,bboxUnit,alt,isPlaceholder,placeholders",
  maxChars: 80000,
});
await fs.writeFile("/Users/susielim/Documents/ChatGPT/YOGA Hosting/.codex-build/exco-deck/reference-inspection.ndjson", snapshot.ndjson);
console.log(snapshot.ndjson);
