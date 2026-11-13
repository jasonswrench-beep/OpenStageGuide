import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {PARTS} from './export-contract.mjs';
import {FACILITATOR_LABEL} from '../../src/facilitator-output.mjs';
export async function auditFacilitator(blob,{forbidden=[],expectedRows,stamp}={}){
 const zip=await JSZip.loadAsync(await blob.arrayBuffer(),{checkCRC32:true}),docs=new Map(),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
 assert.deepEqual(Object.values(zip.files).filter(f=>!f.dir).map(f=>f.name).sort(),PARTS);
 for(const name of PARTS){const raw=await zip.file(name).async('string'),errors=[],doc=new DOMParser({errorHandler:{warning:e=>errors.push(e),error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(raw,'application/xml');assert.deepEqual(errors,[]);assert.doesNotMatch(raw,/<!DOCTYPE|<!ENTITY/);const decoded=raw+doc.documentElement.textContent+Array.from(doc.getElementsByTagName('*')).flatMap(n=>Array.from(n.attributes||[],a=>a.value)).join('');for(const secret of forbidden)assert.ok(!decoded.includes(secret),`Forbidden data in ${name}: ${secret}`);docs.set(name,doc);}
 const R='http://schemas.openxmlformats.org/package/2006/relationships',O='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
 for(const [part,expected] of [['_rels/.rels',[['rId1',O+'officeDocument','word/document.xml'],['rId2','http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties','docProps/core.xml']]],['word/_rels/document.xml.rels',[['rId1',O+'styles','styles.xml']]]]){
  const nodes=Array.from(docs.get(part).getElementsByTagNameNS(R,'Relationship'));assert.deepEqual(nodes.map(n=>[n.getAttribute('Id'),n.getAttribute('Type'),n.getAttribute('Target')]),expected);assert.ok(nodes.every(n=>!n.hasAttribute('TargetMode')));
 }
 const types=docs.get('[Content_Types].xml'),CT='http://schemas.openxmlformats.org/package/2006/content-types';assert.equal(types.getElementsByTagNameNS(CT,'Default').length,2);assert.equal(types.getElementsByTagNameNS(CT,'Override').length,3);
 const core=docs.get('docProps/core.xml').documentElement;assert.deepEqual(Array.from(core.childNodes).filter(n=>n.nodeType===1).map(n=>n.nodeName),['dc:title','dc:creator','dcterms:created']);assert.equal(core.getElementsByTagName('dc:title')[0].textContent,FACILITATOR_LABEL);assert.equal(core.getElementsByTagName('dc:creator')[0].textContent,'OpenStage Guide');if(stamp)assert.equal(core.getElementsByTagName('dcterms:created')[0].textContent,stamp);
 const doc=docs.get('word/document.xml');for(const tag of ['sdt','vanish','webHidden','commentReference','ins','del','object','altChunk','drawing','pict','pBdr','instrText','fldSimple'])assert.equal(doc.getElementsByTagNameNS(W,tag).length,0);
 const walk=n=>n.nodeType===1?(n.namespaceURI===W&&n.localName==='t'?n.textContent:n.namespaceURI===W&&n.localName==='br'?'\n':Array.from(n.childNodes).map(walk).join('')):'';
 const rows=Array.from(doc.getElementsByTagNameNS(W,'p')).map(p=>[p.getElementsByTagNameNS(W,'pStyle')[0]?.getAttributeNS(W,'val'),walk(p)]);
 if(expectedRows)assert.deepEqual(rows,expectedRows);return {rows,text:rows.map(([,text])=>text).join('\n'),zip,docs};
}
