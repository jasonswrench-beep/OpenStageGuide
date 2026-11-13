import JSZip from 'jszip';
import {DOMParser} from '@xmldom/xmldom';
import {fixture,records} from './mock-office.mjs';
import {xml} from '../../src/worksheet.mjs';
export function installParser(){globalThis.DOMParser=DOMParser;}
export async function masterFixture(project=fixture()){
  const zip=new JSZip(),w='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  zip.file('word/document.xml',`<w:document xmlns:w="${w}"><w:body>${records(project).map(r=>`<w:sdt><w:sdtPr><w:tag w:val="${r.tag}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>${xml(r.text).replace(/\n/g,'</w:t><w:br/><w:t>')}</w:t></w:r></w:p></w:sdtContent></w:sdt>`).join('')}<w:p><w:r><w:rPr><w:b/><w:color w:val="663399"/></w:rPr><w:t>UNTAGGED MANUAL NOTE</w:t></w:r></w:p></w:body></w:document>`);
  zip.file('word/webextensions/webextension1.xml',`<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11"><we:properties><we:property name="openstage.guide.project.v1" value="${xml(JSON.stringify(project))}"/><we:property name="unknown.extension.setting" value="KEEP UNKNOWN DATA"/></we:properties></we:webextension>`);
  zip.file('word/media/image1.png',new Uint8Array([137,80,78,71,1,2,3]));
  zip.file('customXml/item1.xml','<unknown>KEEP UNKNOWN METADATA</unknown>');
  zip.file('word/styles.xml','<styles>KEEP ORIGINAL STYLE DEFINITIONS</styles>');
  zip.file('word/settings.xml','<settings>KEEP WORD SETTINGS</settings>');
  return {project,master:await zip.generateAsync({type:'uint8array'})};
}
export const assets=async()=>({ok:true,arrayBuffer:async()=>new Uint8Array([137,80,78,71,4,5]).buffer});
