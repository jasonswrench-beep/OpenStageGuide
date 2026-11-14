// What is checked about a learner workbook beyond whether it leaks (that is workbook-leak.mjs): the structure rules the Review screen
// uses, run on the finished file; a print check for the chosen look; and a look at a PDF for the markers of a tagged one.
// None of this is Word's own Accessibility Checker or a screen-reader test, and the screens say so.
import {checkParts,contrast} from './guide-check.mjs';

const hex6=c=>String(c).replace('#','').toUpperCase();
const grayOf=hex=>{const [r,g,b]=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)),v=Math.round(0.299*r+0.587*g+0.114*b).toString(16).padStart(2,'0');return `${v}${v}${v}`.toUpperCase();};

// The structure rules, on the files of a workbook ({name,text} entries). Returns the errors and warnings only.
export function workbookStructure(files,{DOMParserClass=globalThis.DOMParser}={}){
  const text=name=>files.find(f=>f.name===name)?.text;
  const parse=t=>t?new DOMParserClass().parseFromString(t,'application/xml'):null;
  const core=text('docProps/core.xml')||'',settings=text('word/settings.xml')||'';
  const title=/<dc:title>([^<]*)<\/dc:title>/.exec(core)?.[1]||'',language=/<w:themeFontLang [^>]*w:val="([^"]+)"/.exec(settings)?.[1]||'';
  const {findings,counts}=checkParts({root:parse(text('word/document.xml')),stylesRoot:parse(text('word/styles.xml')),title,language});
  const keep=findings.filter(f=>f.level==='fail'||f.level==='warn');
  return {fix:keep.filter(f=>f.level==='fail'),look:keep.filter(f=>f.level==='warn'),counts};
}

// For the look that was chosen (what resolveWorkbookBrand gave back as `effective`). Each finding is {level,message}.
export function printCheck(effective,{lines=3}={}){
  const gray=effective.printMode==='grayscale',page=hex6(effective.background),text=hex6(effective.text),accent=gray?'333333':hex6(effective.accent);
  const out=[],add=(level,message)=>out.push({level,message});
  const inGrey=(label,fg)=>{
    const ratio=contrast(grayOf(fg),grayOf(page));
    ratio>=4.5?add('pass',`${label} prints as a gray at ${ratio.toFixed(1)} to 1 against the page in black and white.`):add('warn',`${label} prints as a gray that is too light in black and white (${ratio.toFixed(1)} to 1; it needs 4.5 to 1).`);
  };
  inGrey('Body text',text);inGrey('The accent color',accent);
  effective.bodyPt>=11?add('pass',`Body text is ${effective.bodyPt} pt.`):add('warn',`Body text is ${effective.bodyPt} pt, which is small for a printed page.`);
  effective.marginIn>=0.5?add('pass',`Margins are ${effective.marginIn} in, inside what most printers can reach.`):add('warn',`Margins are ${effective.marginIn} in, which some printers will cut off.`);
  lines>0&&effective.responseLineIn<0.25?add('warn','The lines to write on are closer together than 0.25 in.'):lines>0&&add('pass',`Lines to write on are about ${effective.responseLineIn} in apart.`);
  add('info',gray?'This is the black-and-white version. The icons are still colored pictures, and each has its label beside it.':'This is the color version. For a black-and-white copy, choose Black and white. The icons are colored pictures, and each has its label beside it.');
  return out;
}

// What a person still has to check in a workbook, whatever the automatic checks said.
export const WORKBOOK_HUMAN_CHECKS=[
  'Run Word\'s own check on the workbook: Review > Check Accessibility.',
  'Read one page of the workbook with a screen reader before you share it.',
  'Print one page in color and one in black and white, and check that nothing is understood by color alone.'
];

// A PDF is tagged when its catalogue says it is marked and points to a structure tree. This reads the file's plain text, which
// is enough for a PDF whose catalogue is stored in the open, as Word's is; one that stores it compressed is reported as unknown.
export function checkPdfTags(bytes){
  const bin=typeof bytes==='string'?bytes:Array.from(bytes.slice(0,Math.min(bytes.length,8*1024*1024)),b=>String.fromCharCode(b)).join('');
  if(!bin.startsWith('%PDF-'))return {tagged:'no',message:'This is not a PDF file.',has:{}};
  const has={structTree:/\/StructTreeRoot\b/.test(bin),marked:/\/Marked\s+true\b/.test(bin),lang:/\/Lang\s*\(/.test(bin),title:/\/Title\s*[(<]/.test(bin)};
  if(has.structTree&&has.marked)return {tagged:'yes',has,message:`The PDF is tagged${has.lang?' and has a language':', but no language is set'}${has.title?' and a title':', but no title is set'}. Read it with a screen reader to see how it reads.`};
  if(/\/ObjStm\b/.test(bin)&&!has.structTree)return {tagged:'unknown',has,message:'This PDF keeps its catalogue compressed, so the add-in could not tell whether it is tagged. Open it in a desktop PDF program and check its tags.'};
  return {tagged:'no',has,message:'Word did not tag this PDF, so a screen reader cannot read its structure. In Word choose File > Export > Create PDF/XPS > Options > Document structure tags for accessibility, then save the PDF from there.'};
}
