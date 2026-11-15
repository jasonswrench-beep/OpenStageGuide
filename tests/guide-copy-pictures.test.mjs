// A picture copied from one guide into another comes with its picture (replaceParagraphs takes the pictures), under an id the other guide does not use for anything else,
// and a picture already in the guide is shared, not doubled.
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {createDemoAdapter} from '../src/guide-word.mjs';
import {para,run,makeContext,pictureRun} from '../src/guide-ooxml.mjs';
import {pictureSize} from '../src/image-prepare.mjs';

const demo=()=>createDemoAdapter({DOMParserClass:DOMParser,SerializerClass:XMLSerializer});
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const OTHER='R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
const pictureParagraph=(alt,data=PNG,mime='image/png',type='png')=>ctx=>para('OSGCueBody',pictureRun({base64:data,type,mime,size:pictureSize({width:300,height:200},'small'),alt,decorative:false},ctx));

test('a paragraph with a picture is pasted into another guide with the picture, and the other guide\'s own pictures are untouched',async()=>{
  const a=demo(),b=demo();
  await a.insert(pictureParagraph('From the first guide'),{where:'start',spare:false});
  await b.insert(pictureParagraph('Already in the second guide',OTHER,'image/gif','gif'),{where:'start',spare:false});
  const pa=a.exportState(),pb=b.exportState();
  const rid=/r:embed="(rIdPic[A-Za-z0-9]+)"/.exec(pa.paragraphs[0].xml)[1];
  const ownIds=Object.keys(pb.media);
  assert.equal(ownIds.length,1);
  assert.equal(b.mediaOf(rid),null,'the second guide does not have the first one\'s picture');
  await b.replaceParagraphs({start:1,count:0,guard:'',xmls:[pa.paragraphs[0].xml],select:1,pictures:[{rid,...a.mediaOf(rid)}]});
  const after=b.exportState();
  const ids=[...new Set(after.paragraphs.flatMap(p=>[...p.xml.matchAll(/r:embed="(rIdPic[A-Za-z0-9]+)"/g)].map(m=>m[1])))];
  assert.equal(ids.length,2,'two pictures in two paragraphs, each with its own id');
  assert.ok(ids.every(id=>after.media[id]),'and the guide holds both');
  assert.equal(after.media[ownIds[0]].base64,OTHER,'its own picture is as it was');
  const copied=after.paragraphs[1].xml.match(/r:embed="(rIdPic[A-Za-z0-9]+)"/)[1];
  assert.equal(after.media[copied].base64,PNG,'the pasted paragraph names the copy');
});

test('pasting a picture the guide already has shares it, and a picture of a kind Word does not hold is refused with the guide unchanged',async()=>{
  const a=demo();
  await a.insert(pictureParagraph('Mine'),{where:'start',spare:false});
  const state=a.exportState(),rid=/r:embed="(rIdPic[A-Za-z0-9]+)"/.exec(state.paragraphs[0].xml)[1];
  await a.replaceParagraphs({start:1,count:0,guard:'',xmls:[state.paragraphs[0].xml],select:1});
  assert.equal(Object.keys(a.exportState().media).length,1,'one picture, used by two paragraphs');
  assert.equal(a.exportState().paragraphs.filter(p=>p.xml.includes(`r:embed="${rid}"`)).length,2);
  await assert.rejects(a.replaceParagraphs({start:2,count:0,guard:'',xmls:[state.paragraphs[0].xml],pictures:[{rid,mime:'image/svg+xml',base64:PNG}]}),/./);
  assert.equal(a.exportState().paragraphs.length,2,'nothing was added');
});
