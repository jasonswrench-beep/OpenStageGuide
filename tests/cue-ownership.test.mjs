import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import {ICONS} from '../src/model.mjs';
import {CUE_ASSETS} from '../src/cue-assets.mjs';
import {createCueOwnership,cueOwnerTag,inspectOwnedCuePackage} from '../src/cue-ownership.mjs';
import {inspectMaster,sha256,createCheckpoint,readCheckpoint} from '../src/checkpoint.mjs';
import {masterFixture,installParser,assets} from './helpers/checkpoint-fixture.mjs';
import {xml} from '../src/worksheet.mjs';
installParser();
async function ownedFixture(cell=false){
  const {project}=await masterFixture(),item=project.modules[0].activities[0];project.cueOwnership={version:1};item.cueImage=createCueOwnership(item.id,item.icon);item.privateUnknown={preserve:'UNKNOWN ACTIVITY DATA'};
  const {master}=await masterFixture(project),zip=await JSZip.loadAsync(master),png=new Uint8Array(await fs.readFile(new URL('../public/assets/icons/'+CUE_ASSETS[item.icon].file,import.meta.url)));
  const original=await zip.file('word/document.xml').async('string'),anchor=`<w:sdt><w:sdtPr><w:tag w:val="osg:activity:${item.id}"`,at=original.indexOf(anchor),end=original.indexOf('</w:sdt>',at)+8,title=original.slice(at,end);
  const picture=`<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="Instructional cue" title="${xml('OpenStage Guide '+item.icon+' cue')}" descr="${xml(item.icon+' icon')}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="ownedImage"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  const owned=`<w:p><w:sdt><w:sdtPr><w:tag w:val="${cueOwnerTag(item.cueImage)}"/></w:sdtPr><w:sdtContent><w:r>${picture}</w:r></w:sdtContent></w:sdt></w:p>`,cue=`<w:p><w:r><w:t>${xml(item.icon+' · '+item.minutes+' min')}</w:t></w:r></w:p>`,block=owned+cue+title;
  zip.file('word/document.xml',original.replace(title,cell?`<w:tbl><w:tr><w:tc>${block}</w:tc></w:tr></w:tbl>`:block));
  zip.file('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="ownedImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/owned.png"/></Relationships>');zip.file('word/media/owned.png',png);
  zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  return {project,item,zip,picture,owned,cue,title};
}
test('prospective cue tokens are bounded hidden identities and known asset hashes match all supplied PNGs',async()=>{
  assert.deepEqual(Object.keys(CUE_ASSETS),ICONS.map(([label])=>label));
  for(const [label,file] of ICONS){assert.equal(CUE_ASSETS[label].file,file);assert.ok(Object.isFrozen(CUE_ASSETS[label]));assert.equal(CUE_ASSETS[label].sha256,await sha256(await fs.readFile(new URL('../public/assets/icons/'+file,import.meta.url))));}
  const {project}=await masterFixture(),id=project.modules[0].activities[0].id,a=createCueOwnership(id,'Learning activity'),b=createCueOwnership(id,'Learning activity');assert.notEqual(cueOwnerTag(a),cueOwnerTag(b));assert.ok(cueOwnerTag(a).length<=64);assert.equal(a.activityId,id);assert.throws(()=>createCueOwnership('legacy id','Learning activity'));assert.throws(()=>createCueOwnership(id,'Unknown cue'));
});
test('legacy picture masters and unrelated activity metadata are not silently adopted or rewritten',async()=>{
  const {project}=await masterFixture();project.modules[0].activities[0].cueImage={unrelated:'OLD UNKNOWN METADATA'};const {master}=await masterFixture(project),before=await sha256(master),zip=await JSZip.loadAsync(master);
  assert.deepEqual(await inspectOwnedCuePackage(project,zip),{version:1,owned:[],legacy:true,nativeVerified:false});assert.deepEqual((await inspectMaster(master,JSZip)).project,project);assert.equal(await sha256(master),before);
});
test('captured owned-cue proof verifies known image bytes, accessible labels, geometry and same body or cell placement',async()=>{
  for(const cell of [false,true]){
    const f=await ownedFixture(cell),before=JSON.stringify(f.project),data=await f.zip.generateAsync({type:'uint8array'}),proof=await inspectOwnedCuePackage(f.project,f.zip);assert.equal(proof.nativeVerified,false);assert.equal(proof.owned.length,1);assert.equal(proof.owned[0].container,cell?'cell':'body');assert.equal(proof.owned[0].widthPoints,72);assert.equal(proof.owned[0].heightPoints,72);
    assert.deepEqual((await inspectMaster(data,JSZip)).project,f.project);const backup=await createCheckpoint(data,f.project,JSZip,assets);assert.deepEqual((await readCheckpoint(backup.data,JSZip)).master,data);assert.equal(JSON.stringify(f.project),before);assert.deepEqual(await f.zip.file('word/media/image1.png').async('uint8array'),new Uint8Array([137,80,78,71,1,2,3]));
  }
});
test('owned namespace rejects missing duplicate orphan nested moved or altered proof without adopting manual pictures',async()=>{
  const defects=['missing','duplicate','orphan','other-story','nested','moved','extra-picture','manual-image','label','alt','geometry','external','duplicate-rel','child-table','extra-text','other-cell','metadata-hash','unknown-version','duplicate-tag'];
  for(const defect of defects){
    const f=await ownedFixture(),path='word/document.xml';let doc=await f.zip.file(path).async('string'),rels=await f.zip.file('word/_rels/document.xml.rels').async('string');
    if(defect==='missing')doc=doc.replace(f.owned,'');
    if(defect==='duplicate')doc=doc.replace(f.owned,f.owned+f.owned);
    if(defect==='orphan')doc=doc.replace(f.owned,f.owned+f.owned.replace(f.item.cueImage.nonce,'0'.repeat(32)));
    if(defect==='other-story'){doc=doc.replace(f.owned,'');f.zip.file('word/header1.xml',`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${f.owned}</w:hdr>`);}
    if(defect==='nested')doc=doc.replace(f.owned,`<w:sdt><w:sdtPr><w:tag w:val="manual"/></w:sdtPr><w:sdtContent>${f.owned}</w:sdtContent></w:sdt>`);
    if(defect==='moved')doc=doc.replace(f.cue,'<w:p><w:r><w:t>MANUAL GAP</w:t></w:r></w:p>'+f.cue);
    if(defect==='extra-picture')doc=doc.replace(f.picture,f.picture+f.picture);
    if(defect==='manual-image')f.zip.file('word/media/owned.png',new Uint8Array([137,80,78,71,9]));
    if(defect==='label')doc=doc.replace(f.cue,f.cue.replace('10 min','99 min'));
    if(defect==='alt')doc=doc.replace('descr="Learning activity icon"','descr="Manual picture"');
    if(defect==='geometry')doc=doc.replace('cx="914400"','cx="0"');
    if(defect==='external')rels=rels.replace('Target="media/owned.png"','Target="media/owned.png" TargetMode="External"');
    if(defect==='duplicate-rel')rels=rels.replace('</Relationships>',rels.slice(rels.indexOf('<Relationship Id='),rels.indexOf('/>')+2)+'</Relationships>');
    if(defect==='child-table')doc=doc.replace(f.picture,'<w:tbl/>'+f.picture);
    if(defect==='extra-text')doc=doc.replace(f.picture,'<w:t>MANUAL CONTENT</w:t>'+f.picture);
    if(defect==='other-cell')doc=doc.replace(f.owned,`<w:tbl><w:tr><w:tc>${f.owned}</w:tc></w:tr></w:tbl>`);
    if(defect==='metadata-hash')f.item.cueImage.assetHash='0'.repeat(64);
    if(defect==='unknown-version')f.project.cueOwnership.version=2;
    if(defect==='duplicate-tag')doc=doc.replace(`w:val="${cueOwnerTag(f.item.cueImage)}"/>`,`w:val="${cueOwnerTag(f.item.cueImage)}"/><w:tag w:val="${cueOwnerTag(f.item.cueImage)}"/>`);
    f.zip.file(path,doc);f.zip.file('word/_rels/document.xml.rels',rels);await assert.rejects(inspectOwnedCuePackage(f.project,f.zip),undefined,defect);
  }
});
test('owned markers require explicit format opt-in and cannot hide in unsupported package stories',async()=>{
  const f=await ownedFixture();delete f.project.cueOwnership;await assert.rejects(inspectOwnedCuePackage(f.project,f.zip));
  const g=await ownedFixture();g.zip.file('word/unexpectedStory.xml',`<w:doc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${g.owned}</w:doc>`);await assert.rejects(inspectOwnedCuePackage(g.project,g.zip));
});
test('owned picture and timing paragraph profiles reject VML, embedded objects and other unsupported content',async()=>{
  const additions=[
    '<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape><v:imagedata/></v:shape></w:pict>',
    '<w:object xmlns:o="urn:schemas-microsoft-com:office:office"><o:OLEObject/></w:object>',
    '<w:r><w:tab/></w:r>',
    '<w:r><w:rPr><w:vanish/></w:rPr></w:r>'
  ];
  for(const where of ['picture','timing'])for(const extra of additions){
    const f=await ownedFixture(),part='word/document.xml',doc=await f.zip.file(part).async('string'),original=where==='picture'?f.owned:f.cue;
    f.zip.file(part,doc.replace(original,original.replace('</w:p>',extra+'</w:p>')));await assert.rejects(inspectOwnedCuePackage(f.project,f.zip),undefined,where+' '+extra);
  }
});
test('story discovery uses content types and relationships across relocated, uppercase and non-XML filenames',async()=>{
  for(const [path,target,typed,related] of [['misc/header.xml','../misc/header.xml',true,true],['word/header1.XML','header1.XML',true,true],['misc/header.payload','../misc/header.payload',true,false],['misc/header.bin','../misc/header.bin',false,true]]){
    const f=await ownedFixture();f.zip.file(path,`<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${f.owned}</w:hdr>`);
    if(typed){const type=await f.zip.file('[Content_Types].xml').async('string');f.zip.file('[Content_Types].xml',type.replace('</Types>',`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`));}
    if(related){const rels=await f.zip.file('word/_rels/document.xml.rels').async('string');f.zip.file('word/_rels/document.xml.rels',rels.replace('</Relationships>',`<Relationship Id="relocatedHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${target}"/></Relationships>`));}
    await assert.rejects(inspectOwnedCuePackage(f.project,f.zip),undefined,path);
  }
});
