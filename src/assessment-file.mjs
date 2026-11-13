import {validateAssessmentDefinition,ASSESSMENT_LIMITS} from './assessment-scoring.mjs';

// Explicit PRIVATE authoring files only. No storage, responses, scores, consent,
// Word writes, fetching, migrations or automatic transmission is supplied here.
export const ASSESSMENT_FILE_FORMAT='openstage-private-assessment';
export const ASSESSMENT_FILE_VERSION=1;
export const ASSESSMENT_FILE_LIMITS=Object.freeze({maxFileBytes:4*1024*1024,maxDraftTextUnits:1000000,maxJsonDepth:16,maxJsonNodes:100000});
const ITEM_FIELDS=Object.freeze(['item-type','item-prompt','item-target','item-left','item-right','item-high','anchor-preset','anchor-count','anchor-min']);
const RULE_FIELDS=Object.freeze(['rule-name','rule-method','rule-missing','rule-k']);
const AGREEMENT=Object.freeze(['Strongly disagree','Disagree','Neither agree nor disagree','Agree','Strongly agree']);
const fail=m=>{throw new Error(`Unsupported PRIVATE assessment file: ${m}`);};
const arity=(n,wanted)=>{if(n!==wanted)fail('argument count changed.');};
function record(value,keys,label){
 if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be plain own data.`);
 const own=Reflect.ownKeys(value),out=Object.create(null);
 if(own.length!==keys.length||own.some(k=>typeof k!=='string'||!keys.includes(k)))fail(`${label} fields changed.`);
 for(const k of keys){const d=Object.getOwnPropertyDescriptor(value,k);if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${k} must be visible own data.`);out[k]=d.value;}
 return out;
}
function list(value,max,label,min=0){
 if(!Array.isArray(value)||Object.getPrototypeOf(value)!==Array.prototype)fail(`${label} must be a dense array.`);
 const n=Object.getOwnPropertyDescriptor(value,'length')?.value;
 if(!Number.isSafeInteger(n)||n<min||n>max)fail(`${label} array bounds changed.`);
 const keys=Reflect.ownKeys(value);
 if(keys.length!==n+1||keys.some(k=>k!=='length'&&(typeof k!=='string'||!/^(0|[1-9][0-9]*)$/u.test(k)||Number(k)>=n)))fail(`${label} array bounds or fields changed.`);
 const out=[];for(let i=0;i<n;i++){const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}[${i}] must be own data.`);out.push(d.value);}return out;
}
function freeze(value){if(value&&typeof value==='object'){for(const v of Object.values(value))freeze(v);Object.freeze(value);}return value;}
function scalar(value,label,budget,max=16000){
 if(typeof value!=='string'||value.length>max)fail(`${label} text is too large or unsupported.`);
 for(let i=0;i<value.length;i++){const c=value.charCodeAt(i);if(c>=0xd800&&c<=0xdbff){const next=value.charCodeAt(++i);if(!(next>=0xdc00&&next<=0xdfff))fail(`${label} has an unpaired surrogate.`);}else if(c>=0xdc00&&c<=0xdfff)fail(`${label} has an unpaired surrogate.`);}
 budget.units+=value.length;if(budget.units>ASSESSMENT_FILE_LIMITS.maxDraftTextUnits)fail('draft text budget exceeded.');return value;
}
function identifier(v,label,b){if(typeof v!=='string'||!/^[A-Za-z][A-Za-z0-9._:-]{0,79}$/u.test(v))fail(`${label} stable ID changed.`);return scalar(v,label,b,80);}
function choice(value,values,label){if(!values.includes(value))fail(`${label} choice is unavailable.`);return value;}
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
function numeric(v,label,min,max){if(typeof v!=='string'||!/^[-]?(?:0|[1-9][0-9]*)$/u.test(v)||v==='-0')fail(`${label} is unfinished.`);const n=Number(v);if(!Number.isSafeInteger(n)||n<min||n>max)fail(`${label} is out of range.`);return n;}
function countStrings(v,b){if(typeof v==='string')scalar(v,'saved definition text',b);else if(v&&typeof v==='object')for(const child of Object.values(v))countStrings(child,b);}
function itemBuffer(value,b,label){
 const r=record(value,['fields','reverse','anchors'],label),f=record(r.fields,ITEM_FIELDS,`${label}.fields`),fields={};
 for(const k of ITEM_FIELDS)fields[k]=scalar(f[k],`${label}.${k}`,b,k==='anchor-count'?2:k==='anchor-min'?12:16000);
 choice(fields['item-type'],['likert','semantic-differential'],`${label} type`);choice(fields['anchor-preset'],['agreement','frequency','importance','semantic','custom'],`${label} preset`);
 if(typeof r.reverse!=='boolean')fail(`${label} reverse must be boolean.`);
 const seen=new Set(),anchors=list(r.anchors,11,`${label} anchors`,2).map((v,i)=>{const a=record(v,['id','label','code'],`${label} anchor[${i}]`),id=identifier(a.id,`${label} option ID`,b);if(seen.has(id))fail(`${label} option ID duplicated.`);seen.add(id);return {id,label:scalar(a.label,`${label} option label`,b),code:scalar(a.code,`${label} option code`,b,12)};});
 return {fields,reverse:r.reverse,anchors};
}
function ruleBuffer(value,b,known,label){
 const r=record(value,['fields','members'],label),f=record(r.fields,RULE_FIELDS,`${label}.fields`),fields={};
 for(const k of RULE_FIELDS)fields[k]=scalar(f[k],`${label}.${k}`,b,k==='rule-k'?3:16000);
 choice(fields['rule-method'],['mean','sum'],`${label} method`);choice(fields['rule-missing'],['complete-case','minimum-answered'],`${label} missing rule`);
 const seen=new Set(),members=list(r.members,200,`${label} members`).map((v,i)=>{const m=record(v,['id','checked'],`${label} member[${i}]`),id=identifier(m.id,`${label} member ID`,b);if(seen.has(id)||!known.has(id)||typeof m.checked!=='boolean')fail(`${label} membership changed.`);seen.add(id);return {id,checked:m.checked};});
 if(members.length!==known.size)fail(`${label} must retain the full known item roster.`);
 return {fields,members};
}
function compileItem(id,b){
 const f=b.fields;numeric(f['anchor-count'],'anchor count',2,11);numeric(f['anchor-min'],'anchor minimum',-1000000,1000000);
 const item={id,type:f['item-type'],prompt:f['item-prompt'],reverse:b.reverse,options:b.anchors.map(a=>({id:a.id,label:a.label,code:numeric(a.code,'option code',-1000000,1000000)}))};
 if(item.type==='semantic-differential')Object.assign(item,{target:f['item-target'],leftEndpoint:f['item-left'],rightEndpoint:f['item-right'],highScoreMeaning:f['item-high']});
 return validateAssessmentDefinition({schemaVersion:1,id:'a',version:1,title:'x',evidenceStatus:'not-established',items:[item],subscales:[]}).items[0];
}
function compileRule(id,b,items){
 const f=b.fields,rule={id,name:f['rule-name'],itemIds:b.members.filter(m=>m.checked).map(m=>m.id),aggregation:f['rule-method'],missing:{mode:f['rule-missing']}};
 if(rule.missing.mode==='minimum-answered')rule.missing.minimumAnswered=numeric(f['rule-k'],'minimum answered',1,200);
 return validateAssessmentDefinition({schemaVersion:1,id:'a',version:1,title:'x',evidenceStatus:'not-established',items,subscales:[rule]}).subscales[0];
}
function normalizeDraft(value){
 const r=record(value,['definitionId','title','version','items','rules','itemBuffers','ruleBuffers','activeItem','activeRule','itemSelection','ruleSelection'],'draft'),b={units:0};
 const definitionId=identifier(r.definitionId,'assessment ID',b),title=scalar(r.title,'draft title',b),version=scalar(r.version,'draft version',b,12);
 const rawItems=list(r.items,200,'saved items'),rawRules=list(r.rules,50,'saved rules');let items=[],rules=[];
 if(rawItems.length){const validated=validateAssessmentDefinition({schemaVersion:1,id:'a',version:1,title:'x',evidenceStatus:'not-established',items:rawItems,subscales:rawRules});items=validated.items;rules=validated.subscales;}else if(rawRules.length)fail('saved rules require saved items.');
 countStrings(items,b);countStrings(rules,b);const known=new Set(items.map(i=>i.id));
 const itemBuffers=list(r.itemBuffers,200,'saved item buffers').map((v,i)=>{const x=record(v,['id','buffer'],`item buffer[${i}]`),id=identifier(x.id,'saved item buffer ID',b),buffer=itemBuffer(x.buffer,b,`saved item ${id}`);if(items[i]?.id!==id||!same(compileItem(id,buffer),items[i]))fail('saved item buffer does not match its exact saved item.');return {id,buffer};});
 if(itemBuffers.length!==items.length)fail('saved item buffer roster is incomplete.');
 const ruleBuffers=list(r.ruleBuffers,50,'saved rule buffers').map((v,i)=>{const x=record(v,['id','buffer'],`rule buffer[${i}]`),id=identifier(x.id,'saved rule buffer ID',b),buffer=ruleBuffer(x.buffer,b,known,`saved rule ${id}`);if(rules[i]?.id!==id||!same(compileRule(id,buffer,items),rules[i]))fail('saved rule buffer does not match its exact saved rule.');return {id,buffer};});
 if(ruleBuffers.length!==rules.length)fail('saved rule buffer roster is incomplete.');
 const ai=record(r.activeItem,['id','buffer'],'active item'),activeItem={id:identifier(ai.id,'active item ID',b),buffer:itemBuffer(ai.buffer,b,'active item')};
 const ar=record(r.activeRule,['id','buffer'],'active rule'),activeRule={id:identifier(ar.id,'active rule ID',b),buffer:ruleBuffer(ar.buffer,b,known,'active rule')};
 if(activeRule.buffer.members.length!==items.length)fail('active rule must retain the full known item roster.');
 const itemSelection=scalar(r.itemSelection,'item selection',b,80),ruleSelection=scalar(r.ruleSelection,'rule selection',b,80);
 if(itemSelection&&!known.has(itemSelection)||ruleSelection&&!rules.some(s=>s.id===ruleSelection))fail('editor selection refers to an unknown saved row.');
 return {definitionId,title,version,items:structuredClone(items),rules:structuredClone(rules),itemBuffers,ruleBuffers,activeItem,activeRule,itemSelection,ruleSelection};
}
/** Validate a bounded raw authoring draft. Unfinished scalar text stays raw. */
export function validateAssessmentDraft(value){arity(arguments.length,1);return freeze(normalizeDraft(value));}
function editorState(d){
 const ib=d.itemBuffers.find(x=>x.id===d.activeItem.id),rb=d.ruleBuffers.find(x=>x.id===d.activeRule.id),a=d.activeItem.buffer,r=d.activeRule.buffer;
 const pristineItem=!d.items.some(x=>x.id===d.activeItem.id)&&a.reverse===false&&a.fields['item-type']==='likert'&&a.fields['item-prompt']===''&&['item-target','item-left','item-right','item-high'].every(k=>a.fields[k]==='')&&a.fields['anchor-preset']==='agreement'&&a.fields['anchor-count']==='5'&&a.fields['anchor-min']==='1'&&a.anchors.length===5&&a.anchors.every((x,i)=>x.label===AGREEMENT[i]&&x.code===String(i+1));
 const pristineRule=!d.rules.some(x=>x.id===d.activeRule.id)&&r.fields['rule-name']===''&&r.fields['rule-method']==='mean'&&r.fields['rule-missing']==='complete-case'&&r.fields['rule-k']===''&&r.members.length===d.items.length&&r.members.every((m,i)=>m.id===d.items[i].id&&!m.checked);
 return {itemClean:!!ib&&same(a,ib.buffer)||pristineItem,ruleClean:!!rb&&same(r,rb.buffer)||pristineRule};
}
/** Derived under current policy; no serialized dirty/baseline flag is trusted. */
export function assessmentDraftEditorState(value){arity(arguments.length,1);return freeze(editorState(normalizeDraft(value)));}
function compiledBase(d){return validateAssessmentDefinition({schemaVersion:1,id:d.definitionId,version:numeric(d.version,'authored version',1,1000000),title:d.title,evidenceStatus:'not-established',items:d.items,subscales:d.rules});}
/** Saved core identity only, for collision checks. It does not authorize export. */
export function assessmentDraftIdentity(value){arity(arguments.length,1);const d=normalizeDraft(value);try{return compiledBase(d);}catch{return null;}}
/** Definition-only save refuses unfinished active work instead of dropping it. */
export function assessmentDefinitionFromDraft(value){arity(arguments.length,1);const d=normalizeDraft(value),s=editorState(d);if(!s.itemClean||!s.ruleClean)fail('save or explicitly replace unfinished item/rule work before saving a definition-only file.');return compiledBase(d);}
function defaultRuleBuffer(items){return {fields:{'rule-name':'','rule-method':'mean','rule-missing':'complete-case','rule-k':''},members:items.map(x=>({id:x.id,checked:false}))};}
/** Reconstruct editable, current-policy buffers without any tester/review state. */
export function assessmentDraftFromDefinition(value){
 arity(arguments.length,1);const d=validateAssessmentDefinition(value),itemBuffers=d.items.map(item=>({id:item.id,buffer:{fields:{'item-type':item.type,'item-prompt':item.prompt,'item-target':item.target||'','item-left':item.leftEndpoint||'','item-right':item.rightEndpoint||'','item-high':item.highScoreMeaning||'','anchor-preset':'custom','anchor-count':String(item.options.length),'anchor-min':String(Math.min(...item.options.map(o=>o.code)))},reverse:item.reverse,anchors:item.options.map(o=>({id:o.id,label:o.label,code:String(o.code)}))}}));
 const ruleBuffers=d.subscales.map(rule=>({id:rule.id,buffer:{fields:{'rule-name':rule.name,'rule-method':rule.aggregation,'rule-missing':rule.missing.mode,'rule-k':rule.missing.mode==='minimum-answered'?String(rule.missing.minimumAnswered):''},members:[...rule.itemIds.map(id=>({id,checked:true})),...d.items.filter(item=>!rule.itemIds.includes(item.id)).map(item=>({id:item.id,checked:false}))]}}));
 const reserved=new Set([d.id,...d.items.flatMap(i=>[i.id,...i.options.map(o=>o.id)]),...d.subscales.map(s=>s.id)]);let extra='draft-subscale';for(let i=1;reserved.has(extra);i++)extra=`draft-subscale-${i}`;
 return validateAssessmentDraft({definitionId:d.id,title:d.title,version:String(d.version),items:d.items,rules:d.subscales,itemBuffers,ruleBuffers,activeItem:structuredClone(itemBuffers[0]),activeRule:ruleBuffers.length?structuredClone(ruleBuffers[0]):{id:extra,buffer:defaultRuleBuffer(d.items)},itemSelection:d.items[0].id,ruleSelection:d.subscales[0]?.id||''});
}
function normalizeFile(value){const r=record(value,['format','version','kind','payload'],'file');if(r.format!==ASSESSMENT_FILE_FORMAT||r.version!==ASSESSMENT_FILE_VERSION)fail('file format/version is unavailable.');choice(r.kind,['draft','definition'],'file kind');return {format:ASSESSMENT_FILE_FORMAT,version:ASSESSMENT_FILE_VERSION,kind:r.kind,payload:r.kind==='draft'?normalizeDraft(r.payload):validateAssessmentDefinition(r.payload)};}
export function createAssessmentFile(kind,payload){arity(arguments.length,2);return freeze(normalizeFile({format:ASSESSMENT_FILE_FORMAT,version:ASSESSMENT_FILE_VERSION,kind,payload}));}
export function serializeAssessmentFile(value){arity(arguments.length,1);const file=normalizeFile(value),bytes=new TextEncoder().encode(JSON.stringify(file,null,2)+'\n');if(bytes.length>ASSESSMENT_FILE_LIMITS.maxFileBytes)fail('file exceeds the 4 MiB limit.');return bytes;}
function strictJson(text){
 let i=0,nodes=0;const ws=()=>{while(i<text.length&&/[\x20\t\r\n]/u.test(text[i]))i++;};
 function string(key=false){const start=i;if(text[i++]!=='"')fail('invalid JSON string.');while(i<text.length){const c=text.charCodeAt(i++);if(c===34){const token=text.slice(start,i);try{return key?JSON.parse(token):undefined;}catch{fail('invalid JSON string.');}}if(c<32)fail('unescaped JSON control character.');if(c===92){const e=text[i++];if(e==='u'){if(!/^[0-9a-fA-F]{4}$/u.test(text.slice(i,i+4)))fail('invalid JSON escape.');i+=4;}else if(!['"','\\','/','b','f','n','r','t'].includes(e))fail('invalid JSON escape.');}}fail('unterminated JSON string.');}
 function value(depth){if(++nodes>ASSESSMENT_FILE_LIMITS.maxJsonNodes)fail('JSON node budget exceeded.');if(depth>ASSESSMENT_FILE_LIMITS.maxJsonDepth)fail('JSON nesting exceeds this file profile.');ws();const c=text[i];if(c==='{'){i++;ws();const keys=new Set();if(text[i]==='}'){i++;return;}for(;;){ws();const key=string(true);if(keys.has(key))fail('duplicate JSON object key.');keys.add(key);ws();if(text[i++]!==':')fail('invalid JSON object.');value(depth+1);ws();const end=text[i++];if(end==='}')return;if(end!==',')fail('invalid JSON object.');}}else if(c==='['){i++;ws();if(text[i]===']'){i++;return;}for(;;){value(depth+1);ws();const end=text[i++];if(end===']')return;if(end!==',')fail('invalid JSON array.');}}else if(c==='"')string();else{const start=i;while(i<text.length&&!/[\x20\t\r\n,\]}]/u.test(text[i]))i++;const token=text.slice(start,i);if(!/^(?:true|false|null|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)$/u.test(token))fail('invalid JSON value.');}}
 value(0);ws();if(i!==text.length)fail('trailing JSON data.');try{return JSON.parse(text);}catch{fail('invalid JSON.');}
}
const typedLength=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(Uint8Array.prototype),'byteLength').get;
export function parseAssessmentFile(input){
 arity(arguments.length,1);if(!(input instanceof Uint8Array))fail('provide UTF-8 file bytes.');let length;try{length=typedLength.call(input);}catch{fail('file bytes are unavailable.');}if(!length||length>ASSESSMENT_FILE_LIMITS.maxFileBytes)fail('choose a nonempty PRIVATE assessment file up to 4 MiB.');
 const bytes=new Uint8Array(input);if(bytes[0]===0xef&&bytes[1]===0xbb&&bytes[2]===0xbf)fail('UTF-8 BOM is outside this canonical file profile.');let text;try{text=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch{fail('file is not valid UTF-8.');}return freeze(normalizeFile(strictJson(text)));
}
