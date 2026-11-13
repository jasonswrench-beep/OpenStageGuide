import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {course} from '../src/model.mjs';
import {fixture} from './helpers/mock-office.mjs';
import {materialsPreview,materialsParagraphs,materialsFiles,materialsBlob,materialsName} from '../src/materials-output.mjs';
import {auditExport,exportCorpus} from './helpers/export-contract.mjs';
const stamp='2026-10-26T05:00:00.000Z';
function corpus(){const c=exportCorpus();c.privateCanaries.push(c.project.title,...c.project.modules.map(m=>m.objective),...c.project.modules.map(m=>m.title),...c.project.modules.flatMap(m=>m.activities).map(a=>a.participant));c.project.modules.forEach(m=>m.activities.forEach((a,i)=>{a.materials[0].item=`Supply ${i+1} <&> café`;a.materials[0].notes='PRIVATE_MATERIAL_NOTES';a.cueImage={nonce:'PRIVATE_OWNER_NONCE'};}));c.project.pending={journal:'PRIVATE_JOURNAL'};c.privateCanaries.push('PRIVATE_OWNER_NONCE','PRIVATE_JOURNAL','PRIVATE_MATERIAL_NOTES');return c;}
test('materials checklist public projection and all-part contract include saved supplies across all cues without private fields',async()=>{
  const c=corpus(),before=structuredClone(c.project),p=materialsPreview(c.project,stamp),blob=await materialsBlob(c.project,stamp,JSZip);assert.equal(p.rows.length,7);assert.equal(p.rows[0].quantity,5);assert.deepEqual(p.rows[0].activities,['1.1 Activity 1','2.1 Activity 8']);assert.deepEqual(Object.keys(p.rows[0]).sort(),['activities','item','quantity','unit','use']);assert.equal(p.stamp,stamp);await auditExport(blob,{...c,kind:'materials'});assert.deepEqual(c.project,before);
  const zip=await JSZip.loadAsync(await blob.arrayBuffer()),doc=await zip.file('word/document.xml').async('string');assert.ok(doc.includes('Supply 1 &lt;&amp;&gt; café'));assert.equal((await zip.file('docProps/core.xml').async('string')).includes(stamp),true);
});
test('independent checklist audit verifies fractional sums reusable maxima unit case and saved encounter/activity order',async()=>{
  const p=fixture(),a=p.modules[0].activities[0],b={...structuredClone(a),id:crypto.randomUUID(),title:'Second saved activity'};p.modules[0].activities.push(b);const item=(item,quantity,unit,use='consumable')=>({item,quantity,unit,use,notes:'PRIVATE_SUPPLY_NOTE'});
  a.materials=[item('Paper',0.1,'sets'),item('Reusable pens',2.125,'pieces','reusable'),item('Paper',1,'Sets')];b.materials=[item('paper',0.2,'sets'),item('Reusable pens',3.001,'pieces','reusable'),item('Paper',0.001,'sets','reusable')];
  const preview=materialsPreview(p,stamp);assert.deepEqual(preview.rows.map(r=>[r.item,r.quantity,r.unit,r.use]),[['Paper',0.3,'sets','consumable'],['Reusable pens',3.001,'pieces','reusable'],['Paper',1,'Sets','consumable'],['Paper',0.001,'sets','reusable']]);await auditExport(await materialsBlob(p,stamp,JSZip),{project:p,privateCanaries:['PRIVATE_SUPPLY_NOTE','PRIVATE DIRECTIONS','SECRET ANSWER','Learner task.'],kind:'materials'});
});
test('empty and partially configured courses state recorded scope and uncertified completeness without implying no supplies needed',async()=>{
  for(const p of [course('PRIVATE EMPTY COURSE'),fixture()]){const preview=materialsPreview(p,stamp);assert.equal(preview.rows.length,0);assert.equal(preview.configuredActivities,0);assert.match(preview.warning,/may still need materials/);assert.equal(materialsParagraphs(preview).at(-1)[1],'No materials recorded');await auditExport(await materialsBlob(p,stamp,JSZip),{project:p,privateCanaries:[p.title,'PRIVATE DIRECTIONS','SECRET ANSWER','Learner task.'],kind:'materials'});}
});
test('each materials file has a new UUID generic filename and invalid generation stamps are refused',()=>{
  const one=materialsName(),two=materialsName();assert.notEqual(one,two);assert.match(one,/^OpenStage Materials Checklist - [a-f0-9-]{36}\.docx$/);for(const bad of ['PRIVATE_STAMP','2026-02-30T00:00:00.000Z',null])assert.throws(()=>materialsFiles(fixture(),bad));
});
test('materials package audit rejects wrong totals units coverage missing repeated activity labels and extra paragraphs',async()=>{
  const c=corpus(),original=await materialsBlob(c.project,stamp,JSZip);
  for(const defect of ['total','unit','coverage','activity','extra']){const z=await JSZip.loadAsync(await original.arrayBuffer()),path='word/document.xml';let text=await z.file(path).async('string');if(defect==='total')text=text.replace('Quantity: 5','Quantity: 999');if(defect==='unit')text=text.replace('Unit: copies','Unit: Copies');if(defect==='coverage')text=text.replace('14 of 14','0 of 14');if(defect==='activity')text=text.replace('; 2.1 Activity 8','');if(defect==='extra')text=text.replace('</w:body>','<w:p><w:r><w:t>Extra unreviewed paragraph</w:t></w:r></w:p></w:body>');z.file(path,text);await assert.rejects(auditExport(await z.generateAsync({type:'blob'}),{...c,kind:'materials'}),undefined,defect);}
});
test('materials all-part privacy corpus rejects hidden encoded split-run property attachment and relationship leakage',async()=>{
  const c=corpus(),secret=c.privateCanaries[0],original=await materialsBlob(c.project,stamp,JSZip);
  for(const defect of ['hidden','encoded','split','comments','customXML','attachment','property','relationship']){const z=await JSZip.loadAsync(await original.arrayBuffer()),path='word/document.xml',doc=await z.file(path).async('string');
    if(['hidden','encoded','split'].includes(defect)){const payload=defect==='encoded'?[...secret].map(c=>'&#'+c.charCodeAt(0)+';').join(''):defect==='split'?`${secret.slice(0,8)}</w:t></w:r><w:r><w:t>${secret.slice(8)}`:secret;z.file(path,doc.replace('<w:body>',`<w:body><w:p><w:r>${defect==='hidden'?'<w:rPr><w:vanish/></w:rPr>':''}<w:t>${payload}</w:t></w:r></w:p>`));}
    if(defect==='comments')z.file('word/comments.xml',`<comments>${secret}</comments>`);if(defect==='customXML')z.file('customXml/private.xml',`<private>${secret}</private>`);if(defect==='attachment')z.file('word/embeddings/private.bin',new TextEncoder().encode(secret));if(defect==='property'){const path='docProps/core.xml';z.file(path,(await z.file(path).async('string')).replace('</cp:coreProperties>',`<dc:description>${secret}</dc:description></cp:coreProperties>`));}if(defect==='relationship'){const path='word/_rels/document.xml.rels';z.file(path,(await z.file(path).async('string')).replace('</Relationships>','<Relationship Id="private" Type="external" Target="https://example.invalid/secret" TargetMode="External"/></Relationships>'));}
    await assert.rejects(auditExport(await z.generateAsync({type:'blob'}),{...c,kind:'materials'}),undefined,defect);
  }
});

test('materials projection refuses invalid XML noncharacters and lone surrogates in all public text without changing saved source',async()=>{
  for(const invalid of ['\ufffe','\uffff','\ud800','\udfff','\ud800x','x\udfff'])for(const field of ['item','unit','activity']){
    const p=fixture(),a=p.modules[0].activities[0];a.materials=[{item:'Supply',quantity:1,unit:'sets',use:'consumable',notes:'PRIVATE DRAFT'}];if(field==='activity')a.title='Title '+invalid;else a.materials[0][field]='Label '+invalid;const before=structuredClone(p);assert.throws(()=>materialsPreview(p,stamp),/unsupported character/);assert.throws(()=>materialsFiles(p,stamp),/unsupported character/);await assert.rejects(materialsBlob(p,stamp,JSZip),/unsupported character/);assert.deepEqual(p,before);
  }
});
test('valid multilingual emoji tabs line breaks and checkbox text survive the reviewed XML boundary',async()=>{
  const p=fixture(),a=p.modules[0].activities[0];a.title='合作 🙂\r\nتعلّم';a.materials=[{item:'☐ Café 中文 🚀\rSecond line',quantity:1.125,unit:'组\tsets',use:'reusable',notes:'PRIVATE NOTE'}];const before=structuredClone(p),preview=materialsPreview(p,stamp);assert.equal(preview.rows[0].item,a.materials[0].item.replace(/\r\n?/g,'\n'));assert.equal(preview.rows[0].unit,a.materials[0].unit);assert.equal(preview.rows[0].activities[0],'1.1 合作 🙂\nتعلّم');await auditExport(await materialsBlob(p,stamp,JSZip),{project:p,privateCanaries:['PRIVATE NOTE','PRIVATE DIRECTIONS','SECRET ANSWER'],kind:'materials'});assert.deepEqual(p,before);
});
