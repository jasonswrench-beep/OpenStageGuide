import {legacyDocxBlob} from './legacy-docx-package.mjs';
import {planMaterials} from './planning.mjs';
import {xml} from './worksheet.mjs';
import {publicOutputText as text} from './output-text.mjs';
export const MATERIALS_SCOPE='Full course · current saved course order · sequential planning';
export const MATERIALS_WARNING='Recorded requirements only. Completeness is not certified. Activities without entries may still need materials.';
export function materialsPreview(project,stamp=new Date().toISOString()){
  if(typeof stamp!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(stamp)||!Number.isFinite(Date.parse(stamp))||new Date(stamp).toISOString()!==stamp)throw new Error('Use a valid checklist generation stamp.');
  const p=planMaterials(project),labels=new Map();
  project.modules.forEach((m,i)=>m.activities.forEach((a,j)=>{
    for(const material of a.materials||[]){text(material.item);text(material.unit);}
    labels.set(a.id,`${i+1}.${j+1} ${text(a.title).trim()||'(Untitled activity)'}`);
  }));
  // Explicit public projection: do not spread the plan or source records.
  return {projectId:text(p.projectId),revision:p.revision,scope:MATERIALS_SCOPE,stamp,configuredActivities:p.configuredActivities,totalActivities:p.totalActivities,assumption:p.assumption,warning:MATERIALS_WARNING,
    rows:p.rows.map(r=>({item:text(r.item),quantity:r.quantity,unit:text(r.unit),use:r.use,activities:r.activities.map(a=>labels.get(a.id))}))};
}
export function materialsParagraphs(preview){
  const p=preview,rows=[['Title','Materials checklist'],['Normal',p.scopeLabel||p.scope],['Normal',`Project ID: ${p.projectId} · Generation revision: ${p.revision}`],...(p.identityLines||[]).map(line=>['Normal',line]),['Normal',`Generated: ${p.stamp}`],['Normal',`${p.configuredActivities} of ${p.totalActivities} activities have recorded materials · ${p.rows.length} material groups`],['Normal',p.warning],['Normal',p.assumption]];
  if(!p.rows.length)rows.push(['Normal','No materials recorded']);
  for(const r of p.rows){rows.push(['Heading1',`☐ ${r.item}`],['Normal',`Quantity: ${r.quantity} · Unit: ${r.unit} · ${r.use==='consumable'?'Consumed or retained':'Reusable shared pool'}`],['Normal',`Activities: ${r.activities.join('; ')}`]);}
  return rows;
}
export function materialsFiles(project,stamp){
  return materialsPackage(materialsPreview(project,stamp));
}
export function materialsPackage(p){
  const body=materialsParagraphs(p).map(([style,value])=>`<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${value.split('\n').map(t=>`<w:r><w:t xml:space="preserve">${xml(t)}</w:t></w:r>`).join('<w:r><w:br/></w:r>')}</w:p>`).join('');
  return {
    '[Content_Types].xml':'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
    '_rels/.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    'word/_rels/document.xml.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    'word/document.xml':`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:bottom="1080" w:left="1152" w:right="1152"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml':'<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>'+[['Normal',24,null],['Title',40,null],['Heading1',28,0]].map(([name,size,level])=>`<w:style w:type="paragraph" w:styleId="${name}" ${name==='Normal'?'w:default="1"':''}><w:name w:val="${name==='Heading1'?'heading 1':name}"/>${name!=='Normal'?'<w:basedOn w:val="Normal"/>':''}<w:pPr><w:spacing w:before="120" w:after="120"/>${name!=='Normal'?'<w:keepNext/>':''}${level===null?'':`<w:outlineLvl w:val="${level}"/>`}</w:pPr><w:rPr><w:sz w:val="${size}"/></w:rPr></w:style>`).join('')+'</w:styles>',
    'docProps/core.xml':`<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Materials checklist</dc:title><dc:creator>OpenStage Guide</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${p.stamp}</dcterms:created></cp:coreProperties>`
  };
}
export async function materialsBlob(project,stamp,Zip){return legacyDocxBlob(materialsFiles(project,stamp),Zip);}
export const materialsName=()=>`OpenStage Materials Checklist - ${crypto.randomUUID()}.docx`;
