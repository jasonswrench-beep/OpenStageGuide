import {legacyDocxBlob} from './legacy-docx-package.mjs';
import {buildAgenda,reviewCourse} from './planning.mjs';
import {xml} from './worksheet.mjs';
import {publicOutputText as text} from './output-text.mjs';

export const AGENDA_SCOPE='Full course · current planned order';
export function agendaView(project,start='09:00'){
  const agenda=buildAgenda(project,start);
  return {...agenda,title:text(agenda.title),projectId:text(project.id),rows:agenda.rows.map(r=>({...r,title:text(r.title),moduleTitle:text(r.moduleTitle)}))};
}
export function agendaPreview(project,start='09:00') {
  const agenda=agendaView(project,start);
  if(!agenda.rows.length)throw new Error('Add an activity before creating an agenda file.');
  if(![agenda.title,...agenda.rows.flatMap(r=>[r.moduleTitle,r.title])].every(s=>s.trim()))throw new Error('Give the course, modules and activities titles before creating an agenda file.');
  return {...agenda,scope:AGENDA_SCOPE,
    issues:reviewCourse(project).map(({severity,message})=>({severity,message}))};
}
const paragraph=(text,style='Normal')=>`<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${String(text).split('\n').map(t=>`<w:r><w:t xml:space="preserve">${xml(t)}</w:t></w:r>`).join('<w:r><w:br/></w:r>')}</w:p>`;
export function agendaFiles(project,start='09:00',stamp=new Date().toISOString()) {
  return agendaPackage(agendaPreview(project,start),stamp);
}
export function agendaPackage(a,stamp=new Date().toISOString()) {
  stamp=text(stamp);
  let body=paragraph(a.title,'Title')+paragraph(a.subtitle||'Planned workshop agenda','Subtitle')+
    paragraph(a.scopeLabel||a.scope)+paragraph(`Project ID: ${a.projectId} · Generation revision: ${a.revision}`)+
    (a.identityLines||[]).map(line=>paragraph(line)).join('')+
    paragraph(`Generated: ${stamp}`)+paragraph(`${a.minutes} minutes · ${a.start}–${a.end}`);
  let last=null;
  for(const row of a.rows){
    if(row.moduleId!==last){body+=paragraph(row.moduleTitle,'Heading1');last=row.moduleId;}
    body+=paragraph(row.title,'Heading2')+paragraph(`${row.start}–${row.end} · ${row.minutes} minutes`);
  }
  const types='<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'+[['document','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'],['styles','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml']].map(([name,type])=>`<Override PartName="/word/${name}.xml" ContentType="${type}"/>`).join('')+'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>';
  return {
    '[Content_Types].xml':types,
    '_rels/.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    'word/_rels/document.xml.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'word/document.xml':`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:bottom="1080" w:left="1152" w:right="1152"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml':'<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>'+[['Normal',24,null],['Title',44,null],['Subtitle',28,null],['Heading1',32,0],['Heading2',26,1]].map(([name,size,level])=>`<w:style w:type="paragraph" w:styleId="${name}" ${name==='Normal'?'w:default="1"':''}><w:name w:val="${name}"/>${name!=='Normal'?'<w:basedOn w:val="Normal"/>':''}<w:pPr><w:spacing w:before="120" w:after="120"/>${name!=='Normal'?'<w:keepNext/>':''}${level===null?'':`<w:outlineLvl w:val="${level}"/>`}</w:pPr><w:rPr><w:sz w:val="${size}"/></w:rPr></w:style>`).join('')+'</w:styles>',
    'docProps/core.xml':`<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xml(a.title)} Planned agenda</dc:title><dc:creator>OpenStage Guide</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${xml(stamp)}</dcterms:created></cp:coreProperties>`
  };
}
export async function agendaBlob(project,start,Zip) {
  return legacyDocxBlob(agendaFiles(project,start),Zip);
}
export function agendaName(title) {
  return `${title.replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,65)||'Course'} - Agenda - ${globalThis.crypto.randomUUID()}.docx`;
}
