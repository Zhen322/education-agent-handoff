import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const root='/Users/zhenshuai/Desktop/小禾学伴/市场调研-20261029';
const w=await SpreadsheetFile.importXlsx(await FileBlob.load(`${root}/交付/竞品对比表.xlsx`));
const r=await w.inspect({kind:'workbook,sheet,table',maxChars:10000,tableMaxRows:2,tableMaxCols:3});
await fs.writeFile(`${root}/.build/final-workbook-import.txt`,r.ndjson);
for (const [name,range,out] of [['产品概况','G6:J9','prices'],['体验证据','G111:K115','pending'],['体验证据','F52:H54','dates']]) {
 const b=await w.render({sheetName:name,range,scale:1,format:'png'});
 await fs.writeFile(`${root}/.build/previews/check-${out}.png`,new Uint8Array(await b.arrayBuffer()));
}
console.log('Final XLSX imported and selected ranges rendered.');
