import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {pptxFixture} from './helpers/pptx-fixture.mjs';
import {readPptx,pptxDraft,editSlideDraft,MAX_PPTX_BYTES} from '../src/pptx.mjs';
import {fixture} from './helpers/mock-office.mjs';
import {worksheetFiles} from '../src/worksheet.mjs';
globalThis.DOMParser=DOMParser;
for(const id of [256,2147483647,255,2147483648,4294967295])test(`slide ID schema boundary ${id}`,async()=>{
  const bytes=await pptxFixture(async z=>{const path='ppt/presentation.xml',p=await z.file(path).async('string');z.file(path,p.replace('id="258"',`id="${id}"`).replace('id="256" r:id="s1"','id="257" r:id="s1"'));});
  if(id>=256&&id<=2147483647){const s=await readPptx(bytes,'boundary.pptx',JSZip);assert.equal(s.slides[0].sourceId,id);}else await assert.rejects(readPptx(bytes,'boundary.pptx',JSZip),/slide identity/);
});
test('PPTX extraction follows declared source order, retains stable IDs and private notes without source mutation or fetching',async()=>{
  const bytes=await pptxFixture(),before=bytes.slice(),oldFetch=globalThis.fetch;globalThis.fetch=()=>{throw new Error('No fetching');};
  try{const source=await readPptx(bytes,'Sample.pptx',JSZip);assert.deepEqual(bytes,before);assert.deepEqual(source.slides.map(s=>s.sourceId),[258,256]);assert.equal(source.slides[0].title,'Title 2');assert.match(source.slides[1].text,/Text & task 1/);assert.equal(source.slides[1].notes,'PRIVATE SPEAKER NOTE');assert.ok(!JSON.stringify(source).includes('SLIDE NUMBER OMIT'));assert.match(source.source.sha256,/^[a-f0-9]{64}$/);assert.ok(source.warnings.some(w=>/External relationship ignored/.test(w)));assert.ok(Object.isFrozen(source.slides[0]));}finally{globalThis.fetch=oldFetch;}
});
test('review edits and assignments remain separate from frozen baseline and all learner outputs',async()=>{
  const source=await readPptx(await pptxFixture(),'Sample.pptx',JSZip),draft=pptxDraft(source),p=fixture();
  const edited=editSlideDraft(draft,draft.slides[0].id,{title:'Edited',text:'Reviewed slide text',notes:'PRIVATE REVIEW NOTES',moduleId:p.modules[0].id,lesson:'Lesson One'},p.modules.map(m=>m.id));
  assert.equal(source.slides[0].title,'Title 2');assert.equal(draft.slides[0].title,'Title 2');assert.equal(edited.slides[0].lesson,'Lesson One');assert.throws(()=>editSlideDraft(draft,'missing',{},[]));assert.throws(()=>editSlideDraft(draft,draft.slides[0].id,{...edited.slides[0],moduleId:'missing'},[]));
  p.slideReview={source,draft:edited};for(const part of Object.values(worksheetFiles(p)))assert.doesNotMatch(part,/PRIVATE SPEAKER|PRIVATE REVIEW|Reviewed slide text/);
});
test('stable source slide identities survive a changed source order',async()=>{
  const a=await readPptx(await pptxFixture(),'Sample.pptx',JSZip),b=await readPptx(await pptxFixture(async z=>{const p=await z.file('ppt/presentation.xml').async('string');z.file('ppt/presentation.xml',p.replace('<p:sldId id="258" r:id="s2"/><p:sldId id="256" r:id="s1"/>','<p:sldId id="256" r:id="s1"/><p:sldId id="258" r:id="s2"/>'));}),'Sample.pptx',JSZip);assert.equal(a.slides[0].id,b.slides[1].id);assert.notEqual(a.source.sha256,b.source.sha256);
});
test('source extraction preserves internal blank paragraphs, explicit breaks, runs and Unicode',async()=>{
  const bytes=await pptxFixture(async z=>{
    const path='ppt/slides/slide2.xml',xml=await z.file(path).async('string');z.file(path,xml.replace('<a:p><a:r><a:t>Text &amp; task 2</a:t></a:r></a:p>','<a:p><a:r><a:t>Alpha</a:t></a:r><a:r><a:t> — café</a:t></a:r><a:br/><a:r><a:t>東京</a:t></a:r></a:p><a:p/><a:p><a:r><a:t>Omega</a:t></a:r></a:p>'));
    const notes='ppt/notesSlides/notesSlide1.xml',n=await z.file(notes).async('string');z.file(notes,n.replace('<a:t>PRIVATE SPEAKER NOTE</a:t></a:r></a:p>','<a:t>Note one</a:t></a:r></a:p><a:p/><a:p><a:r><a:t>Note three</a:t></a:r></a:p>'));
  });const source=await readPptx(bytes,'paragraphs.pptx',JSZip);assert.match(source.slides[0].text,/Alpha — café\n東京\n\nOmega/);assert.equal(source.slides[1].notes,'Note one\n\nNote three');
});
for(const kind of ['duplicate-id','missing-id','duplicate-target','missing-slide','malformed','entity','macro','embedding','escape','missing-relation'])test('PPTX rejects '+kind,async()=>{
  const bytes=await pptxFixture(async z=>{
    if(kind==='macro')z.file('ppt/vbaProject.bin','active');
    else if(kind==='embedding')z.file('ppt/embeddings/object1.bin','active');
    else if(kind==='missing-slide')z.remove('ppt/slides/slide2.xml');
    else if(kind==='malformed')z.file('ppt/slides/slide2.xml','<broken>');
    else if(kind==='entity')z.file('ppt/slides/slide2.xml','<!DOCTYPE x [<!ENTITY e "bad">]><x/>');
    else if(['duplicate-id','missing-id'].includes(kind)){const s=await z.file('ppt/presentation.xml').async('string');z.file('ppt/presentation.xml',s.replace('id="258"',kind==='duplicate-id'?'id="256"':'id=""'));}
    else{const s=await z.file('ppt/_rels/presentation.xml.rels').async('string');z.file('ppt/_rels/presentation.xml.rels',s.replace('slides/slide2.xml',kind==='duplicate-target'?'slides/slide1.xml':kind==='escape'?'../../escape.xml':'missing.xml'));}
  });await assert.rejects(readPptx(bytes,'Bad.pptx',JSZip));
});
test('PPTX rejects unsafe and duplicate ZIP paths, oversized data and non-PPTX inputs',async()=>{
  await assert.rejects(readPptx(new Uint8Array(MAX_PPTX_BYTES+1),'big.pptx',JSZip));await assert.rejects(readPptx(new Uint8Array(100),'legacy.pptx',JSZip));
  const bytes=await pptxFixture();await assert.rejects(readPptx(bytes,'macro.pptm',JSZip));
  await assert.rejects(readPptx(await pptxFixture(z=>z.file('../unsafe.txt','bad')),'bad.pptx',JSZip),/unsafe/);
  const store=await pptxFixture(()=>{},'STORE'),from=new TextEncoder().encode('ppt/slides/slide2.xml'),to=new TextEncoder().encode('ppt/slides/slide1.xml');
  for(let i=0;i<=store.length-from.length;i++)if(from.every((x,j)=>x===store[i+j]))store.set(to,i);
  await assert.rejects(readPptx(store,'bad.pptx',JSZip),/duplicate package/);
});
test('PPTX rejects forged expanded XML sizes before accumulating oversized data',async()=>{
  const bytes=await pptxFixture(z=>z.file('ppt/slides/slide2.xml',' '.repeat(2*1024*1024+100))),v=new DataView(bytes.buffer);
  for(let at=0;at+46<bytes.length;at++)if(v.getUint32(at,true)===0x02014b50){const n=v.getUint16(at+28,true),name=new TextDecoder().decode(bytes.subarray(at+46,at+46+n));if(name==='ppt/slides/slide2.xml'){v.setUint32(at+24,1024,true);v.setUint32(v.getUint32(at+42,true)+22,1024,true);}}
  await assert.rejects(readPptx(bytes,'forged.pptx',JSZip),/size|length|XML/i);
});
for(const kind of ['default-vba','disguised-ole','embedded-package'])test('rejects active content declared through '+kind,async()=>{
  const bytes=await pptxFixture(async z=>{
    z.file('ppt/data.bin','disguised active content');
    if(kind==='default-vba'){const path='[Content_Types].xml',s=await z.file(path).async('string');z.file(path,s.replace('</Types>','<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'));}
    else{const path='ppt/_rels/presentation.xml.rels',s=await z.file(path).async('string');z.file(path,s.replace('</Relationships>',`<Relationship Id="active1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind==='disguised-ole'?'oleObject':'package'}" Target="data.bin"/></Relationships>`));}
  });await assert.rejects(readPptx(bytes,'disguised.pptx',JSZip),/active|embedded|macro/i);
});
test('encrypted flags and damaged XML CRC are rejected before a preview is offered',async()=>{
  const encrypted=await pptxFixture(),v=new DataView(encrypted.buffer);for(let at=0;at+46<encrypted.length;at++)if(v.getUint32(at,true)===0x02014b50){v.setUint16(at+8,v.getUint16(at+8,true)|1,true);break;}
  await assert.rejects(readPptx(encrypted,'encrypted.pptx',JSZip),/encrypted/);
  const damaged=await pptxFixture(),d=new DataView(damaged.buffer);for(let at=0;at+46<damaged.length;at++)if(d.getUint32(at,true)===0x02014b50){const n=d.getUint16(at+28,true),name=new TextDecoder().decode(damaged.subarray(at+46,at+46+n));if(name==='ppt/slides/slide2.xml'){d.setUint32(at+16,d.getUint32(at+16,true)^1,true);break;}}
  await assert.rejects(readPptx(damaged,'damaged.pptx',JSZip),/damaged XML/);
});
test('duplicate relationship IDs and decks exceeding the slide-count limit are rejected',async()=>{
  const duplicate=await pptxFixture(async z=>{const path='ppt/_rels/presentation.xml.rels',s=await z.file(path).async('string');z.file(path,s.replace('Id="s2"','Id="s1"'));});await assert.rejects(readPptx(duplicate,'duplicate.pptx',JSZip),/relationship identity/);
  const many=await pptxFixture(async z=>{const path='ppt/presentation.xml',s=await z.file(path).async('string'),entries=Array.from({length:501},(_,i)=>`<p:sldId id="${256+i}" r:id="s1"/>`).join('');z.file(path,s.replace(/<p:sldIdLst>.*<\/p:sldIdLst>/,`<p:sldIdLst>${entries}</p:sldIdLst>`));});await assert.rejects(readPptx(many,'many.pptx',JSZip),/500 slides/);
});
