import {validateAssessmentDefinition} from './assessment-scoring.mjs';
import {publicOutputText} from './output-text.mjs';
import {WORKBOOK_BASE_LANGUAGES,FAMILY_WORKBOOK_SETTINGS_XML,FAMILY_WORKBOOK_CONTENT_TYPES_XML,FAMILY_WORKBOOK_DOCUMENT_RELS_XML} from './workbook-family-output.mjs';
import {familyWorkbookBlob} from './workbook-family-flow.mjs';

// Definition-only exports. No response or scoring-result argument is accepted.
export const ASSESSMENT_EXPORT_PROFILE_VERSION=1;
export const ASSESSMENT_EXPORT_LANGUAGES=WORKBOOK_BASE_LANGUAGES;
const fail=message=>{throw new Error(`Unsupported assessment export: ${message}`);};
const own=(v,keys)=>{
  if(!v||typeof v!=='object'||Array.isArray(v)||![Object.prototype,null].includes(Object.getPrototypeOf(v)))fail('options must be an own-data record.');
  const names=Reflect.ownKeys(v);if(names.length!==keys.length||names.some(k=>typeof k!=='string'||!keys.includes(k)))fail('option fields are unsupported.');
  for(const k of keys){const d=Object.getOwnPropertyDescriptor(v,k);if(!d||!d.enumerable||!Object.hasOwn(d,'value'))fail('options must contain enumerable data only.');}
};
const text=(v,label,max=16000)=>{if(typeof v!=='string'||v.length>max)fail(`${label} must be bounded text.`);let clean;try{clean=publicOutputText(v);}catch{fail(`${label} contains unsupported XML text.`);}if(clean!==v.replace(/\r\n?/gu,'\n'))fail(`${label} contains unsupported XML text.`);return clean;};
const xml=s=>s.replace(/[&<>"']/gu,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const freeze=v=>{if(v&&typeof v==='object'){for(const x of Object.values(v))freeze(x);Object.freeze(v);}return v;};
const stamp=v=>{if(typeof v!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/u.test(v)||!Number.isFinite(Date.parse(v))||new Date(v).toISOString()!==v)fail('stamp must be a canonical ISO instant.');return v;};
const paragraph=(v,style='Normal')=>`<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${v.split(/(\n|\t)/u).map(s=>s==='\n'?'<w:r><w:br/></w:r>':s==='\t'?'<w:r><w:tab/></w:r>':`<w:r><w:t xml:space="preserve">${xml(s)}</w:t></w:r>`).join('')}</w:p>`;

export function prepareAssessmentExport(definition,options){
  if(arguments.length!==2)fail('provide a definition and exact export options only.');
  own(options,['route','instructions','baseLanguage','stamp']);
  if(!['respondent','author'].includes(options.route))fail('choose a supported route.');
  if(typeof options.baseLanguage!=='string'||!ASSESSMENT_EXPORT_LANGUAGES.includes(options.baseLanguage))fail('base language is unsupported.');
  const d=validateAssessmentDefinition(definition),route=options.route,instructions=text(options.instructions,'instructions'),generatedAt=stamp(options.stamp),baseLanguage=options.baseLanguage;
  // Projection is constructed field by field; the respondent branch never keeps
  // private codes, keys, option IDs, rules or author-only semantic meaning.
  const identity={assessmentId:text(d.id,'assessment ID'),version:d.version,title:text(d.title,'title')};
  const items=d.items.map(item=>({type:item.type,prompt:text(item.prompt,'prompt'),...(item.type==='semantic-differential'?{target:text(item.target,'target'),leftEndpoint:text(item.leftEndpoint,'left endpoint'),rightEndpoint:text(item.rightEndpoint,'right endpoint')}:{}),labels:item.options.map(o=>text(o.label,'anchor label')),
    ...(route==='author'?{itemId:item.id,reverse:item.reverse,codes:item.options.map(o=>({optionId:o.id,code:o.code})),...(item.type==='semantic-differential'?{highScoreMeaning:text(item.highScoreMeaning,'high keyed-score meaning')}:{})}:{})}));
  const preview={profileVersion:ASSESSMENT_EXPORT_PROFILE_VERSION,route,...identity,generatedAt,baseLanguage,instructions,items,
    ...(route==='author'?{subscales:d.subscales.map(s=>({id:s.id,name:text(s.name,'subscale name'),itemIds:[...s.itemIds],aggregation:s.aggregation,missing:{...s.missing}}))}:{})};
  const rows=[],add=(value,style='Normal')=>rows.push({style,text:value});
  add(identity.title,'Title');add(route==='author'?'PRIVATE author scoring guide':'Blank respondent assessment','Subtitle');
  add(`Assessment ID: ${identity.assessmentId} | Authored version: ${identity.version}`,'Meta');add(`Generated: ${generatedAt} | Export profile: ${ASSESSMENT_EXPORT_PROFILE_VERSION} | Base language: ${baseLanguage}`,'Meta');
  if(route==='author')add('PRIVATE author material. Contains scoring keys and numeric coding. Keep separate from respondent copies. No test responses or scored results are included.');
  else add('Blank response copy. Select at most one listed choice for each item. Leave an item blank when missing; use Not applicable separately when appropriate. No responses are preselected.');
  add('Instructions','Heading1');add(instructions||'Read each prompt and its response labels before choosing one response.');
  if(route==='respondent')add('Author review required before use. Authored prompts, labels and instructions are reproduced as written and may themselves disclose information. Numeric scoring keys are excluded.');
  else {add('Scoring policy','Heading1');add('Validation evidence is not established for this authored measure. This guide supplies raw-score rules only, with no norms, cutoffs, clinical interpretation or high-stakes decisions.');add('Map each selected option to its explicit numeric code. For a reverse-keyed item compute keyed value = L + U - x exactly once, where L and U are that item’s minimum and maximum codes and x is its raw code. Direct items retain their raw code. Display order is independent of numeric coding.');add('Missing and Not applicable are distinct nonnumeric states. Both remain in planned N. A complete-case subscale has no score until every planned item is answered. A minimum-answered mean uses only valid keyed answers after the authored k threshold; report n/N and partial status. Never treat missing or Not applicable as zero. A sum requires complete responses. No overall score is inferred.');}
  add('Items','Heading1');
  for(const [i,item]of items.entries()){
    add(`Item ${i+1}`,'Heading2');add(item.prompt,'Prompt');
    if(item.type==='semantic-differential'){add(`Target: ${item.target}`);add(`Left endpoint: ${item.leftEndpoint}`);add(`Right endpoint: ${item.rightEndpoint}`);}
    if(route==='author'){const values=item.codes.map(o=>o.code),lo=Math.min(...values),hi=Math.max(...values);add(`Item ID: ${item.itemId} | Type: ${item.type} | ${item.reverse?`Reverse keyed: ${lo} + ${hi} - x`:'Direct keyed: retain raw code'} | Range: ${lo} to ${hi}`,'Meta');if(item.highScoreMeaning!==undefined)add(`Intended high keyed-score meaning: ${item.highScoreMeaning}`);}
    for(const [j,label]of item.labels.entries())add(route==='respondent'?`[ ] Position ${j+1} of ${item.labels.length}: ${label}`:`Position ${j+1} of ${item.labels.length}: ${label} | Code: ${item.codes[j].code} | Option ID: ${item.codes[j].optionId}`,'Choice');
    if(route==='respondent')add('[ ] Not applicable','Choice');
  }
  if(route==='author'){add('Explicit subscales','Heading1');if(!preview.subscales.length)add('No aggregate requested. Report item values only.');for(const s of preview.subscales){add(s.name,'Heading2');add(`Subscale ID: ${s.id} | Members in authored order: ${s.itemIds.join(', ')} | Planned N: ${s.itemIds.length}`);add(`Aggregation: ${s.aggregation}. ${s.missing.mode==='complete-case'?'Complete-case: all planned responses required.':`Minimum answered mean: k = ${s.missing.minimumAnswered}; use valid keyed responses only and display n/N with partial status.`} Missing and Not applicable remain in planned N.`);}}
  preview.paragraphs=rows;
  const title=`${identity.title} ${route==='author'?'PRIVATE author scoring guide':'Blank respondent assessment'}`;
  const files={
    '[Content_Types].xml':FAMILY_WORKBOOK_CONTENT_TYPES_XML,
    '_rels/.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    'word/_rels/document.xml.rels':FAMILY_WORKBOOK_DOCUMENT_RELS_XML,
    'word/settings.xml':FAMILY_WORKBOOK_SETTINGS_XML,
    'word/document.xml':`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${rows.map(r=>paragraph(r.text,r.style)).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml':`<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:color w:val="000000"/><w:sz w:val="24"/><w:lang w:val="${baseLanguage}"/></w:rPr></w:rPrDefault></w:docDefaults>${[['Normal',24,null],['Title',40,null],['Subtitle',28,null],['Meta',20,null],['Heading1',30,0],['Heading2',26,1],['Prompt',24,null],['Choice',24,null]].map(([name,size,level])=>`<w:style w:type="paragraph" w:styleId="${name}"${name==='Normal'?' w:default="1"':''}><w:name w:val="${name.startsWith('Heading')?'heading '+name.slice(-1):name}"/>${name!=='Normal'?'<w:basedOn w:val="Normal"/>':''}<w:pPr>${name==='Title'||name==='Subtitle'||name.startsWith('Heading')?'<w:keepNext/>':''}<w:spacing w:before="${name.startsWith('Heading')?180:0}" w:after="100" w:line="288" w:lineRule="auto"/>${name==='Choice'?'<w:ind w:left="240" w:hanging="240"/>':''}${level!==null?`<w:outlineLvl w:val="${level}"/>`:''}</w:pPr><w:rPr>${name.startsWith('Heading')?'<w:b/>':''}<w:color w:val="000000"/><w:sz w:val="${size}"/></w:rPr></w:style>`).join('')}</w:styles>`,
    'docProps/core.xml':`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xml(title)}</dc:title><dc:creator>OpenStage Guide</dc:creator><dc:language>${baseLanguage}</dc:language><dcterms:created xsi:type="dcterms:W3CDTF">${generatedAt}</dcterms:created></cp:coreProperties>`
  };
  return freeze({profileVersion:ASSESSMENT_EXPORT_PROFILE_VERSION,route,preview,files});
}

export function assessmentExportName(route){if(!['respondent','author'].includes(route))fail('filename route is unsupported.');return `OpenStage-${route==='author'?'PRIVATE-Author-Scoring-Guide':'Blank-Respondent-Assessment'}-${globalThis.crypto.randomUUID()}.docx`;}
export async function assessmentExportBlob(definition,options,Zip){if(arguments.length!==3)fail('provide only definition, exact options and ZIP implementation.');const prepared=prepareAssessmentExport(definition,options);return familyWorkbookBlob(prepared.files,Zip);}
