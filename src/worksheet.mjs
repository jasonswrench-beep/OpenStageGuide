import {legacyDocxBlob} from './legacy-docx-package.mjs';
import { participantView,validate } from './model.mjs';
import {reviewCourse} from './planning.mjs';
import {publicOutputText as text} from './output-text.mjs';
import {unsupportedObjective} from './objective-text.mjs';
export const WORKBOOK_SCOPE='Full course · current planned order';
export function worksheetName(title){return `${title.replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,60)||'Course'} - Participant Workbook - ${globalThis.crypto.randomUUID()}.docx`;}
export const xml = s => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
export function workbookPreview(project){
  const v=participantView(project);
  if(project.modules.some(unsupportedObjective))throw new Error('A learning objective has an unsupported value. Preserve this guide and repair that field before creating a participant workbook.');
  return {title:text(v.title),revision:v.revision,projectId:text(project.id),modules:v.modules.map(m=>({title:text(m.title),objective:m.objective?text(m.objective):'',activities:m.activities.map(a=>({title:text(a.title),participant:text(a.participant),lines:a.lines}))}))};
}
function paragraph(text,style='Normal') {
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${String(text).split('\n').map(t=>`<w:r><w:t xml:space="preserve">${xml(t)}</w:t></w:r>`).join('<w:r><w:br/></w:r>')}</w:p>`;
}
export function worksheetFiles(project, stamp=new Date().toISOString()) {
  validate(project);
  if(project.modules.some(unsupportedObjective))throw new Error('A learning objective has an unsupported value. Preserve this guide and repair that field before creating a participant workbook.');
  if(reviewCourse(project).some(x=>x.severity==='error'))throw new Error('Resolve the items marked Needs attention in Review Guide before generating a workbook.');
  return workbookPackage(workbookPreview(project),stamp);
}
export function workbookPackage(v,stamp,scope=WORKBOOK_SCOPE,identityLines=[]){
  stamp=text(stamp);
  let body=paragraph(v.title,'Title')+paragraph('Participant workbook','Subtitle')+paragraph(scope)+paragraph(`Project ID: ${v.projectId} · Generation revision: ${v.revision}`)+identityLines.map(line=>paragraph(line)).join('')+paragraph(`Generated: ${stamp}`);
  for(const m of v.modules) {
    body+=paragraph(m.title,'Heading1');
    if(m.objective) body+=paragraph('Learning objective','Heading2')+paragraph(m.objective);
    for(const a of m.activities) {
      body+=paragraph(a.title,'Heading2')+paragraph(a.participant);
      for(let i=0;i<a.lines;i++) body+='<w:p><w:pPr><w:spacing w:after="140" w:line="360"/><w:pBdr><w:bottom w:val="single" w:sz="3" w:color="BBBBBB"/><w:between w:val="single" w:sz="3" w:color="BBBBBB"/></w:pBdr></w:pPr><w:r><w:t> </w:t></w:r></w:p>';
    }
  }
  return {
    '[Content_Types].xml':'<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
    '_rels/.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    'word/_rels/document.xml.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'word/document.xml':`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:bottom="1080" w:left="1152" w:right="1152"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml':`<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>${[['Normal',24,null],['Title',44,null],['Subtitle',28,null],['Heading1',32,0],['Heading2',26,1]].map(([name,size,level])=>`<w:style w:type="paragraph" w:styleId="${name}" ${name==='Normal'?'w:default="1"':''}><w:name w:val="${name==='Heading1'?'heading 1':name==='Heading2'?'heading 2':name}"/>${name!=='Normal'?'<w:basedOn w:val="Normal"/>':''}<w:pPr><w:spacing w:before="${name==='Normal'?0:180}" w:after="120"/>${name!=='Normal'?'<w:keepNext/>':''}${level!==null?`<w:outlineLvl w:val="${level}"/>`:''}</w:pPr><w:rPr><w:sz w:val="${size}"/><w:color w:val="000000"/>${name.startsWith('Heading')?'<w:b/>':''}</w:rPr></w:style>`).join('')}</w:styles>`,
    'docProps/core.xml':`<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xml(v.title)} Participant workbook</dc:title><dc:creator>OpenStage Guide</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${xml(stamp)}</dcterms:created></cp:coreProperties>`
  };
}
export async function worksheetBlob(project,Zip) {
  return legacyDocxBlob(worksheetFiles(project),Zip);
}
