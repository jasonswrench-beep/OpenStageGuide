// Public authored generated-copy text only; no course/Word/storage/network read.
export const WORKBOOK_RESEARCH_VERSION=1;
export const WORKBOOK_RESEARCH_FIELDS=Object.freeze(['rationale','evidence','application','limitations','references']);
export const WORKBOOK_RESEARCH_SECTIONS=Object.freeze([['rationale','Educational rationale'],['evidence','Research evidence'],['application','Application to this course'],['limitations','Limitations'],['references','References']].map(Object.freeze));
export const WORKBOOK_RESEARCH_VALIDITY=Object.freeze(['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing']);
export const WORKBOOK_RESEARCH_LIMITS=Object.freeze({fieldTextUnits:16000,totalTextUnits:64000});
const freeze=value=>{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value);}return value;};
export const EMPTY_WORKBOOK_RESEARCH=freeze({schemaVersion:1,include:false,...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,'']))});
const fail=message=>{throw new Error(`Opening educational research: ${message}`);};
const own=(value,keys,label)=>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be an own-data record.`);
  const found=Reflect.ownKeys(value);if(found.length!==keys.length||found.some(k=>typeof k!=='string'||!keys.includes(k)))fail(`${label} has unsupported fields.`);
  for(const k of keys){const d=Object.getOwnPropertyDescriptor(value,k);if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${k} must be enumerable data.`);}
};
const text=(value,key)=>{
  if(typeof value!=='string'||value.length>WORKBOOK_RESEARCH_LIMITS.fieldTextUnits)fail(`${key} must be bounded authored text.`);
  for(const symbol of value){const cp=symbol.codePointAt(0);if(!(cp===9||cp===10||cp===13||cp>=0x20&&cp<=0xd7ff||cp>=0xe000&&cp<=0xfffd||cp>=0x10000&&cp<=0x10ffff))fail(`${key} contains unsupported XML text.`);}
  return value.replace(/\r\n?/gu,'\n');
};
export function validateWorkbookResearch(value){
  if(arguments.length!==1)fail('provide one research record only.');
  own(value,['schemaVersion','include',...WORKBOOK_RESEARCH_FIELDS],'research');
  if(value.schemaVersion!==1||typeof value.include!=='boolean')fail('research version or inclusion is unsupported.');
  const fields=Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>[k,text(value[k],k)]));
  if(Object.values(fields).reduce((n,v)=>n+v.length,0)>WORKBOOK_RESEARCH_LIMITS.totalTextUnits)fail('research exceeds64,000 text units.');
  if(value.include&&!Object.values(fields).some(v=>v.trim()))fail('write at least one nonblank authored field before including this section.');
  return freeze({schemaVersion:1,include:value.include,...fields});
}
export function workbookResearchRequest(raw){
  if(arguments.length!==1)fail('provide one raw research record only.');
  own(raw,['schemaVersion','include',...WORKBOOK_RESEARCH_FIELDS],'raw research');
  if(raw.schemaVersion!==1)fail('raw research version is unsupported.');
  own(raw.include,['checked','value','invalid'],'inclusion');own(raw.include.invalid,WORKBOOK_RESEARCH_VALIDITY,'inclusion validity');
  if(typeof raw.include.checked!=='boolean'||typeof raw.include.value!=='string'||raw.include.value.length>100||WORKBOOK_RESEARCH_VALIDITY.some(k=>typeof raw.include.invalid[k]!=='boolean'||raw.include.invalid[k]))fail('the research inclusion control is invalid; its draft was retained.');
  const fields=Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(k=>{
    const field=raw[k];own(field,['value','invalid'],k);own(field.invalid,WORKBOOK_RESEARCH_VALIDITY,`${k} validity`);
    if(WORKBOOK_RESEARCH_VALIDITY.some(flag=>typeof field.invalid[flag]!=='boolean'||raw.include.checked&&field.invalid[flag]))fail(`${k} has unfinished or invalid text; its draft was retained.`);
    return [k,field.value];
  }));
  return validateWorkbookResearch({schemaVersion:1,include:raw.include.checked,...fields});
}
export function workbookResearchRows(value){
  if(arguments.length!==1)fail('provide one research record only.');
  const v=validateWorkbookResearch(value);
  return freeze({schemaVersion:1,include:v.include,heading:'Educational research and rationale',sections:v.include?WORKBOOK_RESEARCH_SECTIONS.filter(([k])=>v[k].trim()).map(([key,heading])=>({key,heading,text:v[key]})):[]});
}
