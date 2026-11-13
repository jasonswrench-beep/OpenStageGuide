import {ICONS} from './model.mjs';
import {normalizeMaterials} from './materials.mjs';
export const MAX_LIBRARY_BYTES=1024*1024;
const FIELDS=['title','icon','minutes','facilitator','participant','answers','lines'];
const fail=()=>{throw new Error('Unsupported activity library. Preserve the original file and use a supported OpenStage library copy.');};
function keys(o,allowed){if(!o||typeof o!=='object'||Array.isArray(o)||Object.keys(o).some(k=>!allowed.includes(k))||allowed.some(k=>!Object.hasOwn(o,k)))fail();}
function text(s,max,required=false){if(typeof s!=='string'||s.length>max||(required&&!s.trim()))fail();return s.replace(/\r\n?/g,'\n').trim();}
export function libraryFields(f){
  keys(f,[...FIELDS,...(Object.hasOwn(f||{},'materials')?['materials']:[])]);
  const out={};for(const k of ['title','icon','facilitator','participant','answers'])out[k]=text(f[k],50000,['title','participant','icon'].includes(k));
  if(!ICONS.some(([label])=>label===out.icon))fail();
  for(const [k,max,integer] of [['minutes',1440,false],['lines',20,true]]){
    if((typeof f[k]!=='number'&&typeof f[k]!=='string')||String(f[k]).trim()==='')fail();
    const n=Number(f[k]);if(!Number.isFinite(n)||n<0||n>max||(integer&&!Number.isInteger(n)))fail();out[k]=n;
  }
  if(Object.hasOwn(f,'materials'))out.materials=normalizeMaterials(f.materials);return out;
}
export function emptyLibrary(){return {format:'openstage-guide-activity-library',version:1,entries:[]};}
export function validateLibrary(data){
  keys(data,['format','version','entries']);if(data.format!=='openstage-guide-activity-library'||data.version!==1||!Array.isArray(data.entries)||data.entries.length>100)fail();
  const ids=new Set(),names=new Set();
  const entries=data.entries.map(e=>{
    keys(e,['id','name','tags','fields']);const id=text(e.id,80,true).toLowerCase(),name=text(e.name,200,true);
    if(!/^[a-f0-9]{8}-(?:[a-f0-9]{4}-){3}[a-f0-9]{12}$/i.test(id)||ids.has(id)||names.has(name.toLowerCase()))fail();ids.add(id);names.add(name.toLowerCase());
    if(!Array.isArray(e.tags)||e.tags.length>20)fail();const tags=e.tags.map(t=>text(t,60,true));if(new Set(tags.map(t=>t.toLowerCase())).size!==tags.length)fail();
    return {id,name,tags,fields:libraryFields(e.fields)};
  });
  const out={format:data.format,version:1,entries};if(new TextEncoder().encode(JSON.stringify(out)).length>MAX_LIBRARY_BYTES)fail();return out;
}
export function readLibrary(textValue){
  if(typeof textValue!=='string'||new TextEncoder().encode(textValue).length>MAX_LIBRARY_BYTES)fail();
  let data;try{data=JSON.parse(textValue);}catch{fail();}return validateLibrary(data);
}
export function saveLibraryCopy(library,name,tags,fields){
  const next=validateLibrary(library);
  next.entries.push({id:crypto.randomUUID(),name,tags,fields:libraryFields(fields)});
  return validateLibrary(next);
}
export function mergeLibrary(library,incoming){
  const a=validateLibrary(library),b=validateLibrary(incoming);return validateLibrary({...a,entries:[...a.entries,...b.entries]});
}
export function libraryPlaceholders(fields){
  const unfinished=s=>/\b(?:TODO|TBD)\b|\[(?:insert|add|replace)\b[^\]]*\]/i.test(s||'');
  return [...['title','facilitator','participant','answers'].filter(k=>unfinished(fields[k])),...(fields.materials||[]).flatMap((m,i)=>['item','unit','notes'].filter(k=>unfinished(m[k])).map(k=>`material ${i+1} ${k}`))];
}
export function libraryName(){return `OpenStage Activity Library - ${crypto.randomUUID()}.json`;}
export function serializeLibrary(library){return JSON.stringify(validateLibrary(library));}
export function renameLibraryCopy(library,id,name,tags,expected){
  const next=validateLibrary(library),entry=next.entries.find(e=>e.id===id);
  if(!entry)throw new Error('This library copy no longer exists. Load another copy.');
  const snapshot=validateLibrary({...emptyLibrary(),entries:[expected]}).entries[0];
  if(JSON.stringify(entry)!==JSON.stringify(snapshot))throw new Error('The library copy changed. Reload its name and tags before saving.');
  entry.name=name;entry.tags=tags;return validateLibrary(next);
}
export function removeLibraryCopy(library,id){
  const next=validateLibrary(library),index=next.entries.findIndex(e=>e.id===id);
  if(index<0)throw new Error('This library copy no longer exists.');
  const [entry]=next.entries.splice(index,1);return {next,removed:{entry,index}};
}
export function undoLibraryRemoval(library,removed){
  const next=validateLibrary(library);keys(removed,['entry','index']);
  if(!Number.isInteger(removed.index)||removed.index<0||removed.index>99)fail();
  const entry=validateLibrary({...emptyLibrary(),entries:[removed.entry]}).entries[0];
  next.entries.splice(Math.min(removed.index,next.entries.length),0,entry);return validateLibrary(next);
}
export function editLibraryContent(library,id,fields,expected){
  const next=validateLibrary(library),entry=next.entries.find(e=>e.id===id);
  if(!entry)throw new Error('This library copy was removed. Your content draft is retained; it cannot recreate the copy.');
  const snapshot=validateLibrary({...emptyLibrary(),entries:[expected]}).entries[0];
  if(JSON.stringify(entry)!==JSON.stringify(snapshot))throw new Error('This library copy changed. Preserve your draft, discard it, then load the current copy.');
  keys(fields,[...FIELDS,...(Object.hasOwn(fields||{},'materials')?['materials']:[])]);
  const before=structuredClone(entry);
  entry.fields=libraryFields({...fields,...(!Object.hasOwn(fields,'materials')&&Object.hasOwn(entry.fields,'materials')?{materials:entry.fields.materials}:{})});
  const checked=validateLibrary(next),after=checked.entries.find(e=>e.id===id);
  return {next:checked,change:JSON.stringify(before)===JSON.stringify(after)?null:{before,after:structuredClone(after)}};
}
export function undoLibraryContent(library,change){
  keys(change,['before','after']);
  const pair=[change.before,change.after].map(entry=>validateLibrary({...emptyLibrary(),entries:[entry]}).entries[0]);
  const [before,after]=pair;if(before.id!==after.id)fail();
  const next=validateLibrary(library),index=next.entries.findIndex(e=>e.id===after.id);
  if(index<0||JSON.stringify(next.entries[index])!==JSON.stringify(after))throw new Error('The edited copy changed or was removed. Undo cannot overwrite newer work.');
  next.entries[index]=before;return validateLibrary(next);
}
