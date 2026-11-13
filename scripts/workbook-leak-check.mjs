// Checks a saved learner workbook (.docx) for anything meant for the facilitator.
//   node scripts/workbook-leak-check.mjs "My workbook.docx" [--source "My guide.docx"]
// With --source, the words of every facilitator paragraph in the guide are searched for in the workbook. Without it, only the
// structure is checked (parts, styles, comments, tracked changes, hidden text, custom content, properties).
// It looks at the saved file only. It is NOT Word's Accessibility Checker and NOT a screen-reader test.
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import fs from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
import {leakCheck} from '../src/workbook-leak.mjs';
import {leakInputsOf} from '../src/guide-workbook.mjs';
import {readStyles} from '../src/guide-check.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const els=n=>Array.from(n.childNodes).filter(c=>c.nodeType===1);

export async function readDocxFiles(bytes){
  const zip=await JSZip.loadAsync(bytes),files=[];
  for(const name of Object.keys(zip.files)){
    if(zip.files[name].dir)continue;
    files.push(/\.(xml|rels)$/.test(name)?{name,text:await zip.file(name).async('string')}:{name,base64:await zip.file(name).async('base64')});
  }
  return files;
}

// What a workbook made from a saved guide must not contain, and what it may say because the guide does: found the same way
// the workbook builder finds them.
export async function leakInputsFromDocx(bytes){
  const zip=await JSZip.loadAsync(bytes);
  const doc=await zip.file('word/document.xml')?.async('string'),styles=await zip.file('word/styles.xml')?.async('string');
  if(!doc)throw new Error('The source is not a Word document: word/document.xml is missing.');
  const parse=t=>new DOMParser().parseFromString(t,'application/xml');
  const root=parse(doc).documentElement,body=els(root).find(n=>n.localName==='body');
  return leakInputsOf({body,styles:readStyles(styles?parse(styles):null)});
}
export const facilitatorTextsFromDocx=async bytes=>(await leakInputsFromDocx(bytes)).facilitatorTexts;

// The workbook is a saved file, so the parts and properties Word adds when it saves are accepted (and still searched).
export async function checkWorkbookFile(workbookBytes,sourceBytes=null){
  const files=await readDocxFiles(workbookBytes);
  const inputs=sourceBytes?await leakInputsFromDocx(sourceBytes):{facilitatorTexts:[],allowedTexts:[]};
  return {...leakCheck({files,...inputs,saved:true},{DOMParserClass:DOMParser}),searchedSource:!!sourceBytes};
}

async function main(argv){
  const args=argv.slice(2),at=args.indexOf('--source');
  const source=at>=0?args.splice(at,2)[1]:null,file=args[0];
  if(!file){console.error('Usage: node scripts/workbook-leak-check.mjs <workbook.docx> [--source <guide.docx>]');return 2;}
  const result=await checkWorkbookFile(await fs.readFile(file),source?await fs.readFile(source):null);
  for(const f of result.findings)console.log(`${f.level.toUpperCase().padEnd(4)} ${f.rule}: ${f.message}`);
  console.log(`\n${result.summary}${result.searchedSource?'':'\n(No --source was given, so the guide\'s facilitator words were not searched for.)'}`);
  return result.passed?0:1;
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href)main(process.argv).then(code=>{process.exitCode=code;}).catch(e=>{console.error(e.message);process.exitCode=2;});
