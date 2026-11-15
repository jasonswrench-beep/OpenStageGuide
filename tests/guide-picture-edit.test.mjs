// Changing a picture that is already in the guide: its description, whether it is decoration, and its width (src/guide-picture.mjs).
import test from 'node:test';
import assert from 'node:assert/strict';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
import {para,pictureRun,makeContext} from '../src/guide-ooxml.mjs';
import {readPicture,setPicture,DESCRIPTION_LIMIT} from '../src/guide-picture.mjs';
import {PICTURE_WIDTHS,pictureSize} from '../src/image-prepare.mjs';
import {STYLE_IDS} from '../src/guide-styles.mjs';

const io={parse:xml=>new DOMParser().parseFromString(xml,'application/xml'),serialize:el=>new XMLSerializer().serializeToString(el)};
const PNG='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
// the paragraph the Insert > Picture dialog makes, with the namespaces an exported paragraph carries
const made=(alt,decorative=false,width='large')=>{
  const ctx=makeContext();
  const xml=para(STYLE_IDS.body,pictureRun({base64:PNG,type:'png',mime:'image/png',size:pictureSize({width:300,height:200},width),alt,decorative},ctx));
  return xml.replace('<w:p>','<w:p xmlns:w="'+W+'" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">');
};

test('a picture is read: its description, whether it is decoration, and which of the program\'s widths it has',()=>{
  const p=readPicture(made('A flip chart'),io);
  assert.equal(p.description,'A flip chart');assert.equal(p.decorative,false);assert.equal(p.width,'large');
  assert.equal(readPicture(made('x',false,'small'),io).width,'small');
  const deco=readPicture(made('',true),io);assert.equal(deco.decorative,true);assert.equal(deco.description,'');
  assert.equal(readPicture('<w:p xmlns:w="'+W+'"><w:r><w:t>words</w:t></w:r></w:p>',io),null,'a paragraph with no picture');
  assert.equal(readPicture('',io),null);assert.equal(readPicture('not xml',io),null);
});

test('the description, the decoration mark and the width change, and nothing else does',()=>{
  const xml=made('A flip chart');
  const out=setPicture(xml,{description:'  Three people at a  whiteboard ',decorative:false,width:'small'},io);
  const p=readPicture(out,io);
  assert.equal(p.description,'Three people at a whiteboard');assert.equal(p.width,'small');
  const small=PICTURE_WIDTHS.find(w=>w.id==='small');
  assert.equal(p.cx,Math.round(small.inches*914400));
  assert.equal(p.cy,Math.round(p.cx*2/3),'the picture keeps its shape (300 by 200)');
  assert.equal((out.match(/<a:ext cx="/g)||[]).length,1);assert.match(out,/<a:ext cx="2011680" cy="1341120"\/>/,'the picture\'s own size agrees with the frame');
  assert.equal(out.match(/r:embed="[^"]+"/)[0],xml.match(/r:embed="[^"]+"/)[0],'it is the same picture');
  assert.equal(setPicture(out,{description:'Three people at a whiteboard',decorative:false,width:'small'},io),out,'the same answers give the same XML');
});

test('a picture can be made decoration and back, and one that is not decoration needs a description',()=>{
  const xml=made('A flip chart');
  const deco=setPicture(xml,{description:'ignored',decorative:true},io);
  const d=readPicture(deco,io);assert.equal(d.decorative,true);assert.equal(d.description,'');
  assert.equal((deco.match(/adec:decorative/g)||[]).length,1,'marked once');
  const back=setPicture(deco,{description:'A flip chart again',decorative:false},io);
  assert.equal(readPicture(back,io).decorative,false);assert.ok(!/adec:decorative/.test(back));
  assert.throws(()=>setPicture(xml,{description:'   ',decorative:false},io),/Describe the picture/);
  assert.throws(()=>setPicture(xml,{description:'x'.repeat(DESCRIPTION_LIMIT+1),decorative:false},io),/description is too long/);
  assert.throws(()=>setPicture(xml,{description:'x',decorative:false,width:'huge'},io),/Choose one of the widths/);
  assert.equal(readPicture(setPicture(deco,{decorative:true,width:'medium'},io),io).decorative,true,'marking it twice does not double the mark');
  assert.equal((setPicture(deco,{decorative:true,width:'medium'},io).match(/adec:decorative/g)||[]).length,1);
});

test('the icons of blocks and paragraphs with more than one picture are not changed here',()=>{
  const icon='<w:p xmlns:w="'+W+'" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:r><w:drawing><wp:inline><wp:extent cx="10" cy="10"/><wp:docPr id="1" name="Icon 1" descr=""/></wp:inline></w:drawing></w:r></w:p>';
  assert.equal(readPicture(icon,io),null);
  assert.throws(()=>setPicture(icon,{description:'x',decorative:false},io),/exactly one picture of yours/);
  const two=made('One').replace('</w:p>',made('Two').replace(/^<w:p[^>]*>/,'').replace('</w:p>','')+'</w:p>');
  assert.equal(readPicture(two,io),null);
});
