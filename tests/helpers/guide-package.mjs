// Helpers for tests that need to look inside the Word content the guide pane builds.
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {readStyles,textOf} from '../../src/guide-check.mjs';
import {esc} from '../../src/guide-ooxml.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const TINY='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

export const parse=x=>{
  const errors=[];
  const d=new DOMParser({errorHandler:{warning:()=>{},error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(x,'application/xml');
  assert.deepEqual(errors,[]);
  return d;
};
export const partXml=(pkg,name)=>new RegExp(`<pkg:part pkg:name="${name}"[^>]*><pkg:xmlData>([\\s\\S]*?)</pkg:xmlData></pkg:part>`).exec(pkg)[1];

// What Word's paragraphs would report for this content: text and style name, in order. The spare paragraph the package
// ends with is not content (Word drops it), so it is left out; spareParagraphs() counts it.
export function paragraphsOf(pkg){
  const doc=parse(partXml(pkg,'/word/document.xml')),styles=readStyles(parse(partXml(pkg,'/word/styles.xml')));
  const all=Array.from(doc.getElementsByTagNameNS(W,'p'));
  if(all.length&&!all.at(-1).childNodes.length)all.pop();
  return all.map(p=>{
    const id=Array.from(p.getElementsByTagNameNS(W,'pStyle'))[0]?.getAttributeNS(W,'val')||'Normal';
    return {text:textOf(p),style:styles.get(id)?.name||id,builtIn:id==='Title'?'Title':undefined};
  });
}

// Turns a flat Office package into a .docx so the structure check can read it.
export async function packageToDocx(pkg,title='Test guide'){
  const zip=new JSZip(),overrides=[];
  for(const [,name,type,inner] of pkg.matchAll(/<pkg:part pkg:name="([^"]+)" pkg:contentType="([^"]+)"[^>]*>([\s\S]*?)<\/pkg:part>/g)){
    const bin=/<pkg:binaryData>([\s\S]*?)<\/pkg:binaryData>/.exec(inner);
    if(bin)zip.file(name.slice(1),Buffer.from(bin[1],'base64'));
    else{
      zip.file(name.slice(1),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${/<pkg:xmlData>([\s\S]*)<\/pkg:xmlData>/.exec(inner)[1]}`);
      if(!name.endsWith('.rels'))overrides.push(`<Override PartName="${name}" ContentType="${type}"/>`);
    }
  }
  zip.file('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${overrides.join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
  zip.file('docProps/core.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(title)}</dc:title></cp:coreProperties>`);
  return zip.generateAsync({type:'nodebuffer'});
}
