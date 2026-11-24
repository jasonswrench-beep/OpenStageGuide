import {readPptx,MAX_PPTX_BYTES} from './pptx.mjs';
export const MAX_REVIEW_BYTES=32*1024*1024;
// Five editable strings, 50,000 UTF-16 units each, six JSON bytes per
// worst-case escaped unit, plus fixed slide identity/keys/punctuation.
export const MAX_EDITOR_BYTES=5*50000*6+256;
const fields=['id','title','text','notes','moduleId','lesson'];
const fail=()=>{throw new Error('Unsupported or mismatched presentation review. Keep the original files; choose the matching PPTX and a supported private review record.');};
function keys(value,names){if(!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).length!==names.length||names.some(k=>!Object.hasOwn(value,k)))fail();}
function string(value,max=50000){if(typeof value!=='string'||value.length>max)fail();}
function list(value,max,check){if(!Array.isArray(value)||value.length>max)fail();value.forEach(check);}
function sourceShape(source){
  keys(source,['format','version','source','warnings','slides']);if(source.format!=='openstage-pptx-preview'||source.version!==1)fail();
  keys(source.source,['name','sha256','bytes']);string(source.source.name,255);if(!source.source.name.toLowerCase().endsWith('.pptx')||!/^[a-f0-9]{64}$/.test(source.source.sha256)||!Number.isSafeInteger(source.source.bytes)||source.source.bytes<22||source.source.bytes>MAX_PPTX_BYTES)fail();
  list(source.warnings,4096,x=>string(x,1000));const ids=new Set();
  list(source.slides,500,(s)=>{
    keys(s,['id','sourceId','sourcePart','order','title','text','notes','warnings']);
    if(!Number.isInteger(s.sourceId)||s.sourceId<256||s.sourceId>2147483647||s.id!==`slide:${s.sourceId}`||ids.has(s.id))fail();ids.add(s.id);
    string(s.sourcePart,65535);for(const k of ['title','text','notes'])string(s[k]);list(s.warnings,20,x=>string(x,1000));
  });
  if(!source.slides.length||source.slides.some((s,i)=>s.order!==i+1))fail();
}
export function validateSlideSource(source){sourceShape(source);return source;}
function draftShape(review,source){
  keys(review,['sourceHash','slides']);if(review.sourceHash!==source.source.sha256||!Array.isArray(review.slides)||review.slides.length!==source.slides.length)fail();
  review.slides.forEach((s,i)=>{keys(s,fields);fields.forEach(k=>string(s[k]));if(s.id!==source.slides[i].id)fail();});
}
function editorShape(editor,source){keys(editor,fields);fields.forEach(k=>string(editor[k]));if(!source.slides.some(s=>s.id===editor.id))fail();}
function shape(record){keys(record,['format','version','source','review','unsavedEditor']);if(record.format!=='openstage-pptx-review'||record.version!==1)fail();sourceShape(record.source);draftShape(record.review,record.source);editorShape(record.unsavedEditor,record.source);}
function same(a,b){if(typeof a!==typeof b||a===null||b===null)return a===b;if(typeof a!=='object')return a===b;if(Array.isArray(a)!==Array.isArray(b))return false;const ak=Object.keys(a).sort(),bk=Object.keys(b).sort();return ak.length===bk.length&&ak.every((k,i)=>k===bk[i]&&same(a[k],b[k]));}
export function serializeReview(source,review,unsavedEditor){
  const record={format:'openstage-pptx-review',version:1,source,review,unsavedEditor};shape(record);
  const text=JSON.stringify(record),encoder=new TextEncoder(),bytes=encoder.encode(text).length,editorBytes=encoder.encode(JSON.stringify(unsavedEditor)).length;
  if(bytes-editorBytes+MAX_EDITOR_BYTES>MAX_REVIEW_BYTES)throw new Error('The private review is too large to reserve safe editing space within its 32 MB limit. Keep the original files and use a smaller presentation.');
  return text;
}
export function parseReviewRecord(text){
  if(typeof text!=='string'||new TextEncoder().encode(text).length>MAX_REVIEW_BYTES)fail();let record;try{record=JSON.parse(text);}catch{fail();}shape(record);
  serializeReview(record.source,record.review,record.unsavedEditor);return record;
}
export async function restoreReview(text,data,name,Zip){
  const record=parseReviewRecord(text);
  const fresh=await readPptx(data,name,Zip),old=record.source;
  // Filename may change while the exact source bytes and parsed baseline stay identical.
  if(fresh.source.sha256!==old.source.sha256||fresh.source.bytes!==old.source.bytes||!same(fresh.slides,old.slides)||!same(fresh.warnings,old.warnings))fail();
  const review=structuredClone(record.review),unsavedEditor=structuredClone(record.unsavedEditor);
  // Keep the archive's display name inside its byte budget. The selected alias
  // is separate UI context, never deck identity or a replacement baseline.
  const source=Object.freeze({...fresh,source:Object.freeze({...fresh.source,name:old.source.name})});
  serializeReview(source,review,unsavedEditor);
  return {source,review,unsavedEditor,selectedName:name};
}
