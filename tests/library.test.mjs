import test from 'node:test';
import assert from 'node:assert/strict';
import {fixture} from './helpers/mock-office.mjs';
import {addActivity} from '../src/model.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
import {emptyLibrary,saveLibraryCopy,mergeLibrary,readLibrary,serializeLibrary,validateLibrary,libraryPlaceholders,renameLibraryCopy,removeLibraryCopy,undoLibraryRemoval,editLibraryContent,undoLibraryContent,MAX_LIBRARY_BYTES} from '../src/library.mjs';
const fields=()=>{const {id,...f}=fixture().modules[0].activities[0];return f;};
test('stored content edits keep identity, name/tags, material shape and independent copies with portable private fields',()=>{
  for(const materials of [undefined,[],[{item:'Board',quantity:2,unit:'sets',use:'reusable',notes:'PRIVATE MATERIAL'}]]){
    const f=fields();if(materials!==undefined)f.materials=materials;
    const l=saveLibraryCopy(emptyLibrary(),'Original',['practice'],f),entry=l.entries[0],before=serializeLibrary(l),loaded=structuredClone(entry.fields),p=fixture(),inserted=addActivity(p,p.modules[0].id,entry.fields);
    const raw={...fields(),title:'Revised',participant:'[insert scenario]',facilitator:'',answers:'PRIVATE REVISED ANSWER',minutes:'0',lines:'0'};
    const {next,change}=editLibraryContent(l,entry.id,raw,entry),updated=next.entries[0];
    assert.equal(updated.id,entry.id);assert.equal(updated.name,entry.name);assert.deepEqual(updated.tags,entry.tags);assert.equal(Object.hasOwn(updated.fields,'materials'),materials!==undefined);assert.deepEqual(updated.fields.materials,materials);
    assert.equal(updated.fields.facilitator,'');assert.equal(updated.fields.lines,0);assert.deepEqual(libraryPlaceholders(updated.fields),['participant']);assert.deepEqual(readLibrary(serializeLibrary(next)),next);
    assert.equal(serializeLibrary(l),before);assert.deepEqual(loaded,entry.fields);assert.equal(inserted.item.title,entry.fields.title);assert.deepEqual(undoLibraryContent(next,change),l);
    const revisedInsertion=addActivity(p,p.modules[0].id,{...updated.fields,participant:'Reviewed learner task'});assert.doesNotMatch(JSON.stringify(worksheetFiles(revisedInsertion.next)),/PRIVATE|SECRET/);assert.notEqual(revisedInsertion.item.id,entry.id);
    assert.equal(editLibraryContent(l,entry.id,fields(),entry).change,null);
  }
});
test('stored content edit refuses full-entry stale/deleted snapshots and preserves unrelated catalog changes',()=>{
  let l=saveLibraryCopy(emptyLibrary(),'One',[],fields());const entry=l.entries[0],other=saveLibraryCopy(l,'Two',[],fields()),raw={...fields(),participant:'Revised private draft'};
  const edited=editLibraryContent(other,entry.id,raw,entry);assert.deepEqual(edited.next.entries[1],other.entries[1]);
  for(const changed of [renameLibraryCopy(l,entry.id,'Renamed',[],entry),edited.next,removeLibraryCopy(l,entry.id).next]){
    const before=serializeLibrary(changed),draft=structuredClone(raw);assert.throws(()=>editLibraryContent(changed,entry.id,raw,entry));assert.equal(serializeLibrary(changed),before);assert.deepEqual(raw,draft);
  }
});
test('content undo preserves unrelated edits and refuses conflicting content/name/removal without consuming caller state',()=>{
  const l=saveLibraryCopy(emptyLibrary(),'One',[],fields()),entry=l.entries[0],{next,change}=editLibraryContent(l,entry.id,{...fields(),title:'Edited'},entry),more=saveLibraryCopy(next,'Two',[],fields());
  const restored=undoLibraryContent(more,change);assert.deepEqual(restored.entries[0],entry);assert.deepEqual(restored.entries[1],more.entries[1]);
  for(const newer of [renameLibraryCopy(next,entry.id,'Changed label',[],next.entries[0]),removeLibraryCopy(next,entry.id).next,editLibraryContent(next,entry.id,{...fields(),title:'Later'},next.entries[0]).next]){const before=serializeLibrary(newer);assert.throws(()=>undoLibraryContent(newer,change));assert.equal(serializeLibrary(newer),before);}
});
test('content updates share the portable compact byte budget and reject malformed raw fields atomically',()=>{
  let l=emptyLibrary();for(let i=0;i<6;i++){const f=fields();f.title='t'.repeat(20000);for(const k of ['facilitator','participant','answers'])f[k]='x'.repeat(50000);l=saveLibraryCopy(l,`Large ${i}`,[],f);}
  const extra=MAX_LIBRARY_BYTES-8-new TextEncoder().encode(serializeLibrary(l)).length;l.entries[0].fields.title+='t'.repeat(extra);l=validateLibrary(l);
  const entry=l.entries[0],raw=structuredClone(entry.fields);raw.title+='é'.repeat(10);const before=serializeLibrary(l);assert.throws(()=>editLibraryContent(l,entry.id,raw,entry));assert.equal(serializeLibrary(l),before);assert.ok(raw.title.endsWith('é'.repeat(10)));
  for(const f of [{...fields(),lines:'1.5'},{...fields(),minutes:''},{...fields(),participant:''},{...fields(),unsupported:[]}])assert.throws(()=>editLibraryContent(l,entry.id,f,entry));
});
test('stored-material changes are bounded private content edits with explicit empty shape and full undo',()=>{
  const l=saveLibraryCopy(emptyLibrary(),'Material copy',[],fields()),entry=l.entries[0],row={item:'Board',quantity:'2.5',unit:'sets',use:'reusable',notes:'PRIVATE MATERIAL EDIT'};
  const edited=editLibraryContent(l,entry.id,{...fields(),materials:[row]},entry);assert.equal(edited.next.entries[0].id,entry.id);assert.equal(edited.next.entries[0].fields.materials[0].quantity,2.5);assert.equal(Object.hasOwn(entry.fields,'materials'),false);assert.deepEqual(undoLibraryContent(edited.next,edited.change),l);
  const snapshot=edited.next.entries[0],removed=editLibraryContent(edited.next,entry.id,{...fields(),materials:[]},snapshot);assert.deepEqual(removed.next.entries[0].fields.materials,[]);assert.deepEqual(undoLibraryContent(removed.next,removed.change),edited.next);assert.deepEqual(readLibrary(serializeLibrary(removed.next)),removed.next);
  const before=serializeLibrary(edited.next);for(const rows of [[{...row,quantity:'bad'}],[row,row],[{...row,quantity:0}],Array.from({length:101},(_,i)=>({...row,item:`Supply ${i}`}))])assert.throws(()=>editLibraryContent(edited.next,entry.id,{...fields(),materials:rows},snapshot));assert.equal(serializeLibrary(edited.next),before);
  const p=fixture(),inserted=addActivity(p,p.modules[0].id,edited.next.entries[0].fields);assert.doesNotMatch(JSON.stringify(worksheetFiles(inserted.next)),/PRIVATE MATERIAL EDIT/);assert.notEqual(inserted.item.id,entry.id);
});
test('library roundtrip retains only supported authoring fields and independent entry identity',()=>{
  const f=fields(),input=structuredClone(f),empty=emptyLibrary(),l=saveLibraryCopy(empty,'Role play',['Teamwork'],f),id=l.entries[0].id;
  assert.deepEqual(empty.entries,[]);assert.deepEqual(f,input);assert.deepEqual(readLibrary(JSON.stringify(l)),l);
  assert.notEqual(id,fixture().modules[0].activities[0].id);assert.match(JSON.stringify(l),/PRIVATE DIRECTIONS/);assert.match(JSON.stringify(l),/SECRET ANSWER/);
  const p=fixture(),a=addActivity(p,p.modules[0].id,l.entries[0].fields),b=addActivity(a.next,p.modules[0].id,l.entries[0].fields);
  assert.notEqual(a.item.id,b.item.id);assert.notEqual(a.item.id,id);l.entries[0].fields.title='Changed source';assert.equal(a.item.title,'Original activity');assert.equal(p.modules[0].activities.length,1);
  assert.doesNotMatch(JSON.stringify(worksheetFiles(b.next)),/PRIVATE DIRECTIONS|SECRET ANSWER/);
});
test('library rejects unexpected project data, unsupported versions, malformed fields and duplicate entries atomically',()=>{
  const l=saveLibraryCopy(emptyLibrary(),'One',['tag'],fields()),before=JSON.stringify(l);
  for(const mutate of [x=>x.version=2,x=>x.project={secret:'no'},x=>x.entries[0].fields.id='source-id',x=>x.entries[0].fields.minutes='',x=>x.entries[0].fields.lines=2.5,x=>x.entries[0].fields.icon='unknown',x=>x.entries[0].tags=['Tag','tag'],x=>x.entries.push(structuredClone(x.entries[0]))]){
    const x=structuredClone(l);mutate(x);assert.throws(()=>validateLibrary(x));
  }
  assert.throws(()=>saveLibraryCopy(l,' one ',[],fields()));assert.throws(()=>mergeLibrary(l,l));assert.equal(JSON.stringify(l),before);
  assert.throws(()=>readLibrary('{bad'));assert.throws(()=>readLibrary(JSON.stringify({...l,__proto__:null,constructor:'unsupported'})));
});
test('library enforces bytes, count and text limits including multibyte input',()=>{
  assert.throws(()=>readLibrary(' '.repeat(MAX_LIBRARY_BYTES+1)));
  assert.throws(()=>readLibrary('é'.repeat(MAX_LIBRARY_BYTES/2+1)));
  const l=emptyLibrary();for(let i=0;i<101;i++)l.entries.push({id:crypto.randomUUID(),name:`Copy ${i}`,tags:[],fields:fields()});assert.throws(()=>validateLibrary(l));
  const f=fields();f.participant='a'.repeat(50001);assert.throws(()=>saveLibraryCopy(emptyLibrary(),'One',[],f));
});
test('library detects unresolved fields and imports independent entries without modifying either source',()=>{
  const f=fields();f.participant='[insert scenario]';f.answers='TBD';assert.deepEqual(libraryPlaceholders(f),['participant','answers']);f.participant='Complete scenario';f.answers='Done';assert.deepEqual(libraryPlaceholders(f),[]);
  const a=saveLibraryCopy(emptyLibrary(),'A',[],f),b=saveLibraryCopy(emptyLibrary(),'B',[],fields()),before=JSON.stringify([a,b]),merged=mergeLibrary(a,b);merged.entries[0].fields.title='Edit';assert.equal(JSON.stringify([a,b]),before);assert.equal(merged.entries.length,2);
});
test('every accepted near-limit library exports and reimports within the same exact byte bound',()=>{
  let l=emptyLibrary();for(let i=0;i<6;i++){const f=fields();f.title='t'.repeat(20000);for(const k of ['facilitator','participant','answers'])f[k]='x'.repeat(50000);l=saveLibraryCopy(l,`Large copy ${i}`,[],f);}
  const extra=MAX_LIBRARY_BYTES-10-new TextEncoder().encode(serializeLibrary(l)).length;l.entries[0].fields.title+='t'.repeat(extra);
  l=validateLibrary(l);const exported=serializeLibrary(l);assert.equal(new TextEncoder().encode(exported).length,MAX_LIBRARY_BYTES-10);assert.deepEqual(readLibrary(exported),l);
  assert.ok(new TextEncoder().encode(JSON.stringify(l,null,2)).length>MAX_LIBRARY_BYTES);
});
test('renaming/tagging preserves activity content, identity and previously inserted copies',()=>{
  const l=saveLibraryCopy(emptyLibrary(),'Original',['old'],fields()),entry=l.entries[0],p=fixture(),inserted=addActivity(p,p.modules[0].id,entry.fields),before=JSON.stringify(l);
  const next=renameLibraryCopy(l,entry.id,'Revised label',['new','practice'],entry);assert.equal(next.entries[0].id,entry.id);assert.deepEqual(next.entries[0].fields,entry.fields);assert.equal(next.entries[0].name,'Revised label');assert.equal(inserted.item.title,entry.fields.title);assert.equal(JSON.stringify(l),before);assert.deepEqual(readLibrary(serializeLibrary(next)),next);
});
test('stale/duplicate/invalid name edits refuse atomically without erasing drafts or source entries',()=>{
  let l=saveLibraryCopy(emptyLibrary(),'One',[],fields());l=saveLibraryCopy(l,'Two',[],fields());const old=l.entries[0],changed=renameLibraryCopy(l,old.id,'Changed',[],old),before=JSON.stringify(changed);
  assert.throws(()=>renameLibraryCopy(changed,old.id,'Another',[],old),/changed/);assert.throws(()=>renameLibraryCopy(l,old.id,' two ',[],old));assert.throws(()=>renameLibraryCopy(l,old.id,'',[],old));assert.throws(()=>renameLibraryCopy(l,'missing','Name',[],old));assert.equal(JSON.stringify(changed),before);
});
test('remove/undo retains original identity/order and never overwrites intervening entries',()=>{
  let l=saveLibraryCopy(emptyLibrary(),'One',[],fields());l=saveLibraryCopy(l,'Two',[],fields());const before=JSON.stringify(l),{next,removed}=removeLibraryCopy(l,l.entries[1].id);assert.equal(next.entries.length,1);assert.equal(JSON.stringify(l),before);assert.deepEqual(undoLibraryRemoval(next,removed),l);
  const intervening=saveLibraryCopy(next,'Three',[],fields()),restored=undoLibraryRemoval(intervening,removed);assert.deepEqual(restored.entries.map(e=>e.name),['One','Two','Three']);
  const conflict=saveLibraryCopy(next,'Two',[],fields()),conflictBefore=serializeLibrary(conflict);assert.throws(()=>undoLibraryRemoval(conflict,removed));assert.equal(serializeLibrary(conflict),conflictBefore);assert.throws(()=>undoLibraryRemoval(l,removed));assert.throws(()=>removeLibraryCopy(l,'missing'));
});
test('empty catalog remains portable and UUID case variants cannot bypass duplicate identity checks',()=>{
  const l=saveLibraryCopy(emptyLibrary(),'One',[],fields()),{next,removed}=removeLibraryCopy(l,l.entries[0].id);assert.deepEqual(readLibrary(serializeLibrary(next)),emptyLibrary());assert.deepEqual(undoLibraryRemoval(next,removed),l);
  const duplicate=structuredClone(l.entries[0]);duplicate.id=duplicate.id.toUpperCase();duplicate.name='Different name';assert.throws(()=>validateLibrary({...l,entries:[l.entries[0],duplicate]}));assert.equal(validateLibrary({...l,entries:[duplicate]}).entries[0].id,l.entries[0].id);
});
