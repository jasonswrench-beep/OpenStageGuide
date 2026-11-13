import {legacyDocxBlob} from './legacy-docx-package.mjs';
import {validate,courseDetails,COURSE_DETAIL_FIELDS,ICONS} from './model.mjs';
import {hasFrontMatter} from './frontmatter.mjs';
import {normalizeMaterials} from './materials.mjs';
import {publicOutputText} from './output-text.mjs';
import {xml} from './worksheet.mjs';
export const FACILITATOR_LABEL='Private facilitator guide — structured copy';
export const FACILITATOR_SCOPE='Full course · current saved course order';
export const FACILITATOR_PRIVATE='PRIVATE: Includes teaching directions and saved preparation notes. Share only with intended facilitators.';
export const FACILITATOR_LIMITS='Structured text only. Unmanaged Word additions, images, formatting/layout, headers/footers, attachments and full-master fidelity are omitted. Keep the original master and full backup. Edits to this copy do not update the master. Delivery variants and coordinated packs are unavailable.';
export function facilitatorOptions(value={answers:false,participant:false}){
 if(!value||typeof value!=='object'||Array.isArray(value)||Reflect.ownKeys(value).length!==2)throw new Error('Choose supported facilitator-copy inclusion options.');
 const answers=Object.getOwnPropertyDescriptor(value,'answers'),participant=Object.getOwnPropertyDescriptor(value,'participant');
 if(!answers||!participant||!Object.hasOwn(answers,'value')||!Object.hasOwn(participant,'value')||typeof answers.value!=='boolean'||typeof participant.value!=='boolean')throw new Error('Choose supported facilitator-copy inclusion options.');
 return {answers:answers.value,participant:participant.value};
}
function text(value,label,required=false){
 if(value===undefined&&!required)return '';
 if(typeof value!=='string'||value.length>50000)throw new Error(`Enter supported ${label} text before creating this private copy.`);
 let result;try{result=publicOutputText(value);}catch{throw new Error(`The selected ${label} contains an unsupported character. Preserve your work and correct that field before creating this private copy.`);}
 if(required&&!result.trim())throw new Error(`Enter ${label} before creating this private copy.`);return result;
}
function generation(stamp){if(typeof stamp!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(stamp)||!Number.isFinite(Date.parse(stamp))||new Date(stamp).toISOString()!==stamp)throw new Error('Use a valid private-copy generation stamp.');return stamp;}
export function facilitatorPreview(project,options,stamp=new Date().toISOString()){
 validate(project);const choice=facilitatorOptions(options),issues=[];
 const unfinished=(value,label)=>{if(/\b(?:TODO|TBD)\b|\[(?:insert|add|replace)\b[^\]]*\]/i.test(value))issues.push(`${label}: contains an unfinished prompt.`);};
 const details=courseDetails(project),v={title:text(project.title,'course title',true),projectId:text(project.id,'project ID',true),revision:project.revision,stamp:generation(stamp),options:choice,scope:FACILITATOR_SCOPE,privacy:FACILITATOR_PRIVATE,limits:FACILITATOR_LIMITS,details:[],frontMatter:[],modules:[],issues};
 unfinished(v.title,'Course title');
 for(const key of COURSE_DETAIL_FIELDS){const value=text(details[key],key);if(value.trim()){v.details.push({label:key==='authors'?'Authors':key[0].toUpperCase()+key.slice(1),text:value});unfinished(value,key);}}
 if(hasFrontMatter(project))for(const section of project.frontMatter.sections){if(!section.body.trim())continue;const heading=text(section.heading,'front-matter heading',true),body=text(section.body,'front-matter body');if(body.trim()){v.frontMatter.push({heading,body});unfinished(heading,'Front-matter heading');unfinished(body,'Front-matter body');}}
 if(!project.modules.length)issues.push('No learning modules are recorded. This copy is not evidence of a complete teaching guide.');
 v.modules=project.modules.map((module,mi)=>{
  const m={title:text(module.title,'module title',true),objective:text(module.objective,'learning objective'),activities:[]};unfinished(m.title,`Module ${mi+1}`);if(!m.objective.trim())issues.push(`Module ${mi+1}: learning objective is blank.`);else unfinished(m.objective,`Module ${mi+1} objective`);
  if(!module.activities.length)issues.push(`Module ${mi+1}: no activities are recorded.`);
  m.activities=module.activities.map((activity,ai)=>{
   const label=`${mi+1}.${ai+1}`;
   if(!ICONS.some(([name])=>name===activity.icon))throw new Error('Choose a supported instructional cue label before creating this private copy.');
   if(!Number.isFinite(activity.minutes)||activity.minutes<0||activity.minutes>1440)throw new Error('Fix the activity duration before creating this private copy.');
   const a={title:text(activity.title,'activity title',true),cue:activity.icon,minutes:activity.minutes,facilitator:text(activity.facilitator,'facilitator instructions'),materials:normalizeMaterials(activity.materials===undefined?[]:activity.materials).map(r=>({item:text(r.item,'material name',true),quantity:r.quantity,unit:text(r.unit,'material unit',true),use:r.use,notes:text(r.notes,'preparation notes')}))};
   unfinished(a.title,`Activity ${label} title`);if(!a.facilitator.trim())issues.push(`Activity ${label}: facilitator instructions are blank.`);else unfinished(a.facilitator,`Activity ${label} facilitator instructions`);if(!a.minutes)issues.push(`Activity ${label}: no teaching time is allocated.`);
   if(choice.answers){a.answers=text(activity.answers,'answers/debrief');if(a.answers.trim())unfinished(a.answers,`Activity ${label} answers/debrief`);}
   if(choice.participant){if(!Number.isInteger(activity.lines)||activity.lines<0||activity.lines>20)throw new Error('Fix participant writing-line settings before creating this private copy.');a.participant=text(activity.participant,'participant instructions');a.lines=activity.lines;if(a.participant.trim())unfinished(a.participant,`Activity ${label} participant instructions`);else issues.push(`Activity ${label}: selected participant instructions are blank.`);}
   for(const r of a.materials){unfinished(r.item,`Activity ${label} material`);unfinished(r.notes,`Activity ${label} preparation notes`);}
   return a;
  });return m;
 });return v;
}
export function facilitatorParagraphs(v){
 const rows=[['Title',v.title],['Subtitle',FACILITATOR_LABEL],['Normal',v.privacy],['Normal',v.limits],['Normal',v.scope],['Normal',`Project ID: ${v.projectId} · Generation revision: ${v.revision}`],...(v.identityLines||[]).map(line=>['Normal',line]),['Normal',`Generated: ${v.stamp}`],['Normal',`Include answers/debrief: ${v.options.answers?'Yes':'No'} · Include participant instructions: ${v.options.participant?'Yes':'No'}`]];
 for(const d of v.details)rows.push(['Heading2',d.label],['Normal',d.text]);
 for(const s of v.frontMatter)rows.push(['Heading1',s.heading],['Normal',s.body]);
 v.modules.forEach((m,mi)=>{rows.push(['Heading1',`${mi+1}. ${m.title}`]);if(m.objective.trim())rows.push(['Heading2','Learning objective'],['Normal',m.objective]);m.activities.forEach((a,ai)=>{
  rows.push(['Heading2',`${mi+1}.${ai+1} ${a.title}`],['Normal',`${a.cue} · ${a.minutes} minutes`]);
  if(a.facilitator.trim())rows.push(['Heading3','Facilitator instructions'],['Normal',a.facilitator]);
  if(v.options.answers&&a.answers.trim())rows.push(['Heading3','Answers / debrief'],['Normal',a.answers]);
  if(v.options.participant){if(a.participant.trim())rows.push(['Heading3','Participant instructions'],['Normal',a.participant]);rows.push(['Normal',`Participant writing-line setting: ${a.lines}. No response-space pages are generated in this facilitator copy.`]);}
  if(a.materials.length){rows.push(['Heading3','Saved activity materials']);for(const r of a.materials){rows.push(['Normal',`${r.item} · ${r.quantity} ${r.unit} · ${r.use==='consumable'?'Consumed or retained':'Reusable shared pool'}`]);if(r.notes.trim())rows.push(['Normal',`Preparation notes: ${r.notes}`]);}}
 });});
 if(v.issues.length){rows.push(['Heading1','Unresolved authoring suggestions']);for(const issue of v.issues)rows.push(['Normal',issue]);}
 return rows;
}
export function facilitatorFiles(project,options,stamp){
 return facilitatorPackage(facilitatorPreview(project,options,stamp));
}
export function facilitatorPackage(v){
 const body=facilitatorParagraphs(v).map(([style,value])=>`<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${value.split('\n').map(t=>`<w:r><w:t xml:space="preserve">${xml(t)}</w:t></w:r>`).join('<w:r><w:br/></w:r>')}</w:p>`).join('');
 return {
  '[Content_Types].xml':'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
  '_rels/.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
  'word/_rels/document.xml.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  'word/document.xml':`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:bottom="1080" w:left="1152" w:right="1152"/></w:sectPr></w:body></w:document>`,
  'word/styles.xml':'<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>'+[['Normal',24,null],['Title',40,null],['Subtitle',28,null],['Heading1',32,0],['Heading2',28,1],['Heading3',24,2]].map(([name,size,level])=>`<w:style w:type="paragraph" w:styleId="${name}" ${name==='Normal'?'w:default="1"':''}><w:name w:val="${name}"/>${name!=='Normal'?'<w:basedOn w:val="Normal"/>':''}<w:pPr><w:spacing w:before="120" w:after="120"/>${name!=='Normal'?'<w:keepNext/>':''}${level===null?'':`<w:outlineLvl w:val="${level}"/>`}</w:pPr><w:rPr><w:sz w:val="${size}"/></w:rPr></w:style>`).join('')+'</w:styles>',
  'docProps/core.xml':`<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${FACILITATOR_LABEL}</dc:title><dc:creator>OpenStage Guide</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${v.stamp}</dcterms:created></cp:coreProperties>`
 };
}
export async function facilitatorBlob(project,options,stamp,Zip){return legacyDocxBlob(facilitatorFiles(project,options,stamp),Zip);}
export const facilitatorName=()=>`OpenStage Private Facilitator Guide - ${crypto.randomUUID()}.docx`;
