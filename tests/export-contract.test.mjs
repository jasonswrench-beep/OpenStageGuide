import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {worksheetBlob} from '../src/worksheet.mjs';
import {agendaBlob} from '../src/agenda.mjs';
import {auditExport,exportCorpus} from './helpers/export-contract.mjs';

for(const optionalBlanks of [false,true])test(`combined export contract covers all 14 current cues, two modules and optional blanks=${optionalBlanks}`,async()=>{
  const corpus=exportCorpus({optionalBlanks}),before=JSON.stringify(corpus.project);
  for(const kind of ['workbook','agenda']){
    const blob=kind==='workbook'?await worksheetBlob(corpus.project,JSZip):await agendaBlob(corpus.project,'23:55',JSZip);
    assert.deepEqual(await auditExport(blob,{...corpus,kind,start:'23:55'}),{kind,parts:6,privateCanariesChecked:corpus.privateCanaries.length,activities:14});
  }
  assert.equal(JSON.stringify(corpus.project),before,'Exports must not mutate the master');
});

test('export audit rejects missing repeated learner lines and corrupted row/overall agenda timing',async()=>{
  const corpus=exportCorpus();
  for(const defect of ['missing-repeated-lines','row-time','total-time']){
    const kind=defect==='missing-repeated-lines'?'workbook':'agenda',blob=kind==='workbook'?await worksheetBlob(corpus.project,JSZip):await agendaBlob(corpus.project,'23:55',JSZip);
    const zip=await JSZip.loadAsync(await blob.arrayBuffer()),part=zip.file('word/document.xml');let xml=await part.async('string');
    if(kind==='workbook'){let seen=0;xml=xml.replace(/<w:t xml:space="preserve">Second line Ω<\/w:t>/g,match=>++seen===1?match:'<w:t xml:space="preserve"></w:t>');assert.equal(seen,14);}
    else if(defect==='row-time')xml=xml.replace('23:55–23:55 · 0 minutes','23:56–23:55 · 0 minutes');
    else xml=xml.replace('19.5 minutes · 23:55–Day 2 · 00:14:30','19.5 minutes · 23:55–Day 2 · 00:15:30');
    assert.notEqual(xml,await part.async('string'));zip.file(part.name,xml);
    const changed=await zip.generateAsync({type:'blob'});await assert.rejects(()=>auditExport(changed,{...corpus,kind,start:'23:55'}),undefined,`Audit missed ${defect}`);
  }
});

test('exact export contract rejects duplicate paragraphs, zero-line sections and misplaced content in both outputs',async()=>{
  const corpus=exportCorpus();
  for(const kind of ['workbook','agenda'])for(const defect of ['duplicate-paragraph','duplicate-zero-line-section','misplaced-paragraph']){
    const blob=kind==='workbook'?await worksheetBlob(corpus.project,JSZip):await agendaBlob(corpus.project,'23:55',JSZip),zip=await JSZip.loadAsync(await blob.arrayBuffer()),part=zip.file('word/document.xml'),xml=await part.async('string');
    const paragraphs=xml.match(/<w:p>[\s\S]*?<\/w:p>/g),at=paragraphs.findIndex(p=>p.includes('>Activity 1</w:t>'));assert.ok(at>=0);
    const instruction=paragraphs[at+1],section=paragraphs[at]+instruction;let changed;
    if(defect==='duplicate-paragraph')changed=xml.replace(instruction,instruction+instruction);
    else if(defect==='duplicate-zero-line-section')changed=xml.replace(section,section+section);
    else{const later=paragraphs.find(p=>p.includes('>Activity 2</w:t>'));changed=xml.replace(later,instruction+later);}
    assert.notEqual(changed,xml);zip.file(part.name,changed);
    await assert.rejects(()=>zip.generateAsync({type:'blob'}).then(data=>auditExport(data,{...corpus,kind,start:'23:55'})),undefined,`Audit missed ${kind} ${defect}`);
  }
});

test('independent export audit rejects hidden, comment, customXML, attachment, property and relationship leak controls',async()=>{
  const corpus=exportCorpus(),secret=corpus.privateCanaries[0];
  const mutations={
    hidden:async z=>{const f=z.file('word/document.xml');z.file(f.name,(await f.async('string')).replace('<w:body>',`<w:body><w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>${secret}</w:t></w:r></w:p>`));},
    encodedHidden:async z=>{const f=z.file('word/document.xml');z.file(f.name,(await f.async('string')).replace('<w:body>',`<w:body><w:p><w:r><w:t>${[...secret].map(c=>'&#'+c.charCodeAt(0)+';').join('')}</w:t></w:r></w:p>`));},
    splitRuns:async z=>{const f=z.file('word/document.xml');z.file(f.name,(await f.async('string')).replace('<w:body>',`<w:body><w:p><w:r><w:t>${secret.slice(0,8)}</w:t></w:r><w:r><w:t>${secret.slice(8)}</w:t></w:r></w:p>`));},
    comments:z=>z.file('word/comments.xml',`<comments>${secret}</comments>`),
    customXML:z=>z.file('customXml/item1.xml',`<private>${secret}</private>`),
    attachment:z=>z.file('word/embeddings/private.bin',Buffer.from(secret)),
    property:async z=>{const f=z.file('docProps/core.xml');z.file(f.name,(await f.async('string')).replace('</cp:coreProperties>',`<dc:description>${secret}</dc:description></cp:coreProperties>`));},
    externalRelationship:async z=>{const f=z.file('word/_rels/document.xml.rels');z.file(f.name,(await f.async('string')).replace('</Relationships>',`<Relationship Id="leak" Type="external" Target="https://example.invalid/private" TargetMode="External"/></Relationships>`));}
  };
  for(const kind of ['workbook','agenda']){
    const original=kind==='workbook'?await worksheetBlob(corpus.project,JSZip):await agendaBlob(corpus.project,'09:00',JSZip);
    for(const [name,mutate] of Object.entries(mutations)){
      const zip=await JSZip.loadAsync(await original.arrayBuffer());await mutate(zip);
      const blob=await zip.generateAsync({type:'blob'});
      await assert.rejects(()=>auditExport(blob,{...corpus,kind}),undefined,`Audit missed ${kind} ${name}`);
    }
  }
});
