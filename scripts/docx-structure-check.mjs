// An automated structure check for a .docx. It looks at the saved file only.
// It is NOT Word's Accessibility Checker and NOT a screen-reader test: a pass here is not a native result.
// The rules live in src/guide-check.mjs so the add-in pane and this command use the same ones.
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import fs from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
import {checkParts,readOutParts,contrast} from '../src/guide-check.mjs';

export {contrast};

const parse=(text,name)=>{
  const errors=[];
  const doc=new DOMParser({errorHandler:{warning:()=>{},error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(text,'application/xml');
  if(errors.length)throw new Error(`${name} is not well-formed XML: ${errors[0]}`);
  return doc;
};

async function open(bytes){
  const zip=await JSZip.loadAsync(bytes);
  const text=async n=>{const f=zip.file(n);return f?f.async('string'):null;};
  const docXml=await text('word/document.xml');
  if(!docXml)throw new Error('This is not a Word document: word/document.xml is missing.');
  const stylesXml=await text('word/styles.xml'),core=await text('docProps/core.xml'),settings=await text('word/settings.xml');
  return {
    root:parse(docXml,'word/document.xml'),
    stylesRoot:stylesXml?parse(stylesXml,'word/styles.xml'):null,
    title:/<dc:title>([^<]*)<\/dc:title>/.exec(core||'')?.[1]||'',
    language:/<w:themeFontLang [^>]*w:val="([^"]+)"/.exec(settings||'')?.[1]||''
  };
}

export async function checkDocx(bytes){return checkParts(await open(bytes));}
export async function readOut(bytes){return readOutParts(await open(bytes));}

async function main(){
  const args=process.argv.slice(2),files=args.filter(a=>!a.startsWith('--'));
  if(!files.length){console.error('Usage: node scripts/docx-structure-check.mjs <file.docx> [more files] [--read-out] [--json]');process.exit(2);}
  let failed=false;
  for(const file of files){
    const bytes=await fs.readFile(file),result=await checkDocx(bytes);
    if(args.includes('--json')){console.log(JSON.stringify({file,...result,readOut:args.includes('--read-out')?await readOut(bytes):undefined},null,2));}
    else{
      console.log(`\nStructure check: ${file}`);
      for(const f of result.findings)console.log(`  ${f.level.toUpperCase().padEnd(4)}  ${f.rule.padEnd(16)} ${f.message}`);
      console.log(`  Summary: ${result.counts.fail} fail, ${result.counts.warn} warn, ${result.counts.pass} pass.`);
      if(args.includes('--read-out')){console.log('  Approximate reading order (not a screen reader):');for(const l of await readOut(bytes))console.log(`    ${l}`);}
    }
    if(result.counts.fail)failed=true;
  }
  if(!args.includes('--json'))console.log('\nThis is an automated structure check of the saved file. It is not Word\'s Accessibility Checker or a screen-reader test.');
  process.exit(failed?1:0);
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href)main().catch(e=>{console.error(e.message);process.exit(2);});
