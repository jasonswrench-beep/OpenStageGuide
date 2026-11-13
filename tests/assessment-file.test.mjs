import test from 'node:test';
import assert from 'node:assert/strict';
import {ASSESSMENT_FILE_FORMAT,ASSESSMENT_FILE_VERSION,ASSESSMENT_FILE_LIMITS,validateAssessmentDraft,assessmentDraftFromDefinition,assessmentDraftEditorState,assessmentDraftIdentity,assessmentDefinitionFromDraft,createAssessmentFile,serializeAssessmentFile,parseAssessmentFile} from '../src/assessment-file.mjs';
import {validateAssessmentDefinition,createAssessmentResponse,scoreAssessment} from '../src/assessment-scoring.mjs';
import {assessmentFileDefinition,assessmentFileDraft} from './helpers/assessment-file-fixture.mjs';
const copy=x=>structuredClone(x),bytes=s=>new TextEncoder().encode(s),parsed=x=>parseAssessmentFile(serializeAssessmentFile(x));
function frozen(x){if(x&&typeof x==='object'){assert.ok(Object.isFrozen(x));for(const v of Object.values(x))frozen(v);}}

test('PRIVATE definition files retain both types, authored order/codes/scoring and Unicode without responses or authority',()=>{
 assert.equal(ASSESSMENT_FILE_FORMAT,'openstage-private-assessment');assert.equal(ASSESSMENT_FILE_VERSION,1);
 for(const options of [2,5,11]){const input=assessmentFileDefinition(options,2,-3),before=copy(input),file=createAssessmentFile('definition',input),raw=serializeAssessmentFile(file),read=parseAssessmentFile(raw);assert.deepEqual(read,file);assert.deepEqual(read.payload,validateAssessmentDefinition(before));assert.deepEqual(input,before);assert.equal(Object.isFrozen(input),false);frozen(read);assert.deepEqual(Object.keys(read),['format','version','kind','payload']);assert.ok(!JSON.stringify(read).includes('responses'));
  const entries=read.payload.items.map(i=>({itemId:i.id,state:'answered',optionId:i.options[0].id}));assert.deepEqual(scoreAssessment(read.payload,createAssessmentResponse(read.payload,entries)),scoreAssessment(input,createAssessmentResponse(input,entries)));
  raw.fill(0);assert.equal(read.payload.title,before.title);const second=serializeAssessmentFile(file);assert.deepEqual(parseAssessmentFile(second),file);
 }
});
test('PRIVATE drafts retain unfinished/inactive raw values and derive dirty state instead of trusting saved claims',()=>{
 const draft=assessmentFileDraft({unfinished:true}),before=copy(draft),file=createAssessmentFile('draft',draft),read=parsed(file);assert.deepEqual(read.payload,before);assert.deepEqual(draft,before);frozen(read);assert.deepEqual(assessmentDraftEditorState(read.payload),{itemClean:false,ruleClean:false});assert.equal(assessmentDraftIdentity(read.payload),null);assert.throws(()=>assessmentDefinitionFromDraft(read.payload),/unfinished/);
 draft.activeItem.buffer.fields['item-prompt']='Later';assert.notEqual(read.payload.activeItem.buffer.fields['item-prompt'],'Later');
 const clean=assessmentFileDraft();assert.deepEqual(assessmentDraftEditorState(clean),{itemClean:true,ruleClean:true});assert.deepEqual(assessmentDefinitionFromDraft(clean),assessmentDraftIdentity(clean));
 const unfinished=copy(clean);unfinished.activeItem.buffer.fields['item-prompt']='not yet saved';assert.deepEqual(assessmentDraftIdentity(unfinished),assessmentDraftIdentity(clean));assert.throws(()=>assessmentDefinitionFromDraft(unfinished),/unfinished/);
 for(const field of ['dirty','baseline','serial','validity','responses','consent']){const x=copy(clean);x[field]=false;assert.throws(()=>validateAssessmentDraft(x));}
});
test('loaded definition buffers preserve display and subscale member order and current clean defaults without data coercion',()=>{
 const definition=assessmentFileDefinition(11,3,-3),draft=assessmentDraftFromDefinition(definition);assert.deepEqual(assessmentDefinitionFromDraft(draft),validateAssessmentDefinition(definition));assert.deepEqual(draft.ruleBuffers[0].buffer.members.filter(x=>x.checked).map(x=>x.id),definition.subscales[0].itemIds);
 const d=copy(draft);d.activeRule={id:d.ruleBuffers[1].id,buffer:copy(d.ruleBuffers[1].buffer)};assert.equal(assessmentDraftEditorState(d).ruleClean,true);assert.deepEqual(parsed(createAssessmentFile('draft',d)).payload,d);const subset=copy(d);subset.ruleBuffers[1].buffer.members=subset.ruleBuffers[1].buffer.members.filter(x=>x.checked);assert.throws(()=>validateAssessmentDraft(subset),/full known/);
 const bad=copy(d);bad.activeRule.buffer.members.pop();assert.throws(()=>validateAssessmentDraft(bad),/full known/);
 for(const mutate of [x=>x.itemBuffers[0].buffer.anchors[0].label='Changed saved anchor',x=>x.itemBuffers[0].buffer.fields['anchor-count']='-',x=>x.ruleBuffers[0].buffer.members.reverse(),x=>x.ruleBuffers[0].buffer.members.push({id:'unknown',checked:false}),x=>x.itemBuffers.pop(),x=>x.ruleBuffers.pop()]){const x=copy(draft);mutate(x);assert.throws(()=>validateAssessmentDraft(x));}
 const none={...copy(definition),subscales:[]};const noRules=assessmentDraftFromDefinition(none);assert.equal(assessmentDraftEditorState(noRules).ruleClean,true);assert.deepEqual(assessmentDefinitionFromDraft(noRules),validateAssessmentDefinition(none));
});
test('strict file JSON rejects malformed UTF8, duplicate decoded keys, BOM, trailing data and unsupported nested shapes',()=>{
 const file=createAssessmentFile('draft',assessmentFileDraft()),json=new TextDecoder().decode(serializeAssessmentFile(file));
 for(const s of [json.replace('"version": 1','"version": 1, "version": 1'),json.replace('"kind": "draft"','"kind": "draft", "\\u006bind": "draft"'),json+' false','', '[', '{"x":'+ '['.repeat(10000)+'0'+']'.repeat(10000)+'}',JSON.stringify({...file,version:2}),JSON.stringify({...file,kind:'respondent'}),JSON.stringify({...file,payload:null}),json.replace('"kind": "draft"','"kind": "draft", "responses": []')])assert.throws(()=>parseAssessmentFile(bytes(s)));
 for(const raw of [new Uint8Array([0xc0,0xaf]),new Uint8Array([0xed,0xa0,0xbd]),new Uint8Array([0xf0,0x9f]),Uint8Array.from([0xef,0xbb,0xbf,...bytes(json)])])assert.throws(()=>parseAssessmentFile(raw));
 const lone=copy(file);lone.payload.activeItem.buffer.fields['item-prompt']='\ud800';assert.throws(()=>createAssessmentFile('draft',lone.payload),/surrogate/);assert.throws(()=>parseAssessmentFile(bytes(JSON.stringify(lone))),/surrogate/);
 assert.throws(()=>parseAssessmentFile(bytes('['+'0,'.repeat(100001)+'0]')),/node budget/);
});
test('strict own-data and dense-array checks refuse getters without invocation and retain source descriptors',()=>{
 const draft=assessmentFileDraft(),before=copy(draft),desc=Object.getOwnPropertyDescriptors(draft);let reads=0;
 for(const mutate of [x=>Object.defineProperty(x,'title',{enumerable:true,get(){reads++;throw Error('getter');}}),x=>Object.defineProperty(x.activeItem.buffer.fields,'item-prompt',{enumerable:true,get(){reads++;throw Error('getter');}}),x=>Object.defineProperty(x.items,'0',{enumerable:true,get(){reads++;throw Error('getter');}}),x=>Object.defineProperty(x.activeRule.buffer.members,'0',{enumerable:true,get(){reads++;throw Error('getter');}}),x=>x.items.extra=true,x=>x.items[Symbol.iterator]=()=>{reads++;throw Error('iterator');},x=>delete x.items[0],x=>x[Symbol('hidden')]=1,x=>Object.defineProperty(x,'hidden',{value:1}),x=>Object.setPrototypeOf(x,{secret:true}),x=>Object.defineProperty(x.activeItem.buffer.anchors[0],'code',{value:'1',enumerable:false})]){const x=copy(draft);mutate(x);assert.throws(()=>validateAssessmentDraft(x));}
 assert.equal(reads,0);assert.deepEqual(draft,before);assert.deepEqual(Object.getOwnPropertyDescriptors(draft),desc);
 for(const x of [null,undefined,[],false,0,'draft'])assert.throws(()=>validateAssessmentDraft(x));
});
test('bounded draft/file limits refuse oversize inputs before offering data without silently truncating raw work',()=>{
 const d=assessmentFileDraft({unfinished:true});d.activeItem.buffer.fields['item-prompt']='x'.repeat(16000);assert.equal(validateAssessmentDraft(d).activeItem.buffer.fields['item-prompt'].length,16000);d.activeItem.buffer.fields['item-prompt']+='x';assert.throws(()=>createAssessmentFile('draft',d),/large/);
 const short=assessmentFileDraft();for(const [field,value]of [['version','1'.repeat(13)]]){const x=copy(short);x[field]=value;assert.throws(()=>validateAssessmentDraft(x));}
 for(const field of ['anchor-count','anchor-min']){const x=copy(short);x.activeItem.buffer.fields[field]='1'.repeat(field==='anchor-count'?3:13);assert.throws(()=>validateAssessmentDraft(x));}
 const x=copy(short);x.activeRule.buffer.fields['rule-k']='1234';assert.throws(()=>validateAssessmentDraft(x));
 assert.throws(()=>parseAssessmentFile(new Uint8Array(ASSESSMENT_FILE_LIMITS.maxFileBytes+1)),/4 MiB/);
 const huge=copy(short);huge.items.length=1000000000;assert.throws(()=>validateAssessmentDraft(huge),/bounds/);
 const maxDefinition=assessmentFileDefinition(11,200,0);maxDefinition.subscales=Array.from({length:50},(_,i)=>({...copy(maxDefinition.subscales[0]),id:`scale-${i}`,name:`Scale ${i}`}));const big=assessmentDraftFromDefinition(maxDefinition);const round=parsed(createAssessmentFile('draft',big));assert.equal(round.payload.items.length,200);assert.equal(round.payload.rules.length,50);assert.deepEqual(round.payload,big);const canonical=serializeAssessmentFile(createAssessmentFile('definition',assessmentFileDefinition()));const padded=new Uint8Array(ASSESSMENT_FILE_LIMITS.maxFileBytes);padded.fill(32);padded.set(canonical);assert.equal(parseAssessmentFile(padded).payload.title,assessmentFileDefinition().title);
 const over=copy(short);for(let i=0;i<50;i++){const n=copy(over.itemBuffers[0]);n.id=`item-large-${i}`;n.buffer.fields['item-target']='x'.repeat(16000);n.buffer.fields['item-high']='y'.repeat(16000);const item=copy(over.items[0]);item.id=n.id;over.items.push(item);over.itemBuffers.push(n);}over.activeRule.buffer.members=over.items.map(i=>({id:i.id,checked:false}));assert.throws(()=>createAssessmentFile('draft',over),/budget/);
});
test('byte snapshots and exact function arities preserve caller inputs and never admit responses or restored consent',()=>{
 const file=createAssessmentFile('definition',assessmentFileDefinition()),raw=serializeAssessmentFile(file),buffer=Buffer.from(raw),before=Buffer.from(buffer);Object.defineProperty(buffer,'byteLength',{get(){throw Error('must use native bounds');}});const read=parseAssessmentFile(buffer);buffer.fill(0);assert.deepEqual(read,file);assert.notEqual(buffer[0],before[0]);
 for(const value of [raw.buffer,new DataView(raw.buffer),[],null,undefined,'{}'])assert.throws(()=>parseAssessmentFile(value));
 for(const run of [()=>createAssessmentFile('draft'),()=>createAssessmentFile('definition',file.payload,1),()=>serializeAssessmentFile(file,1),()=>parseAssessmentFile(raw,1),()=>validateAssessmentDraft(assessmentFileDraft(),1),()=>assessmentDraftFromDefinition(file.payload,1),()=>assessmentDraftEditorState(assessmentFileDraft(),1),()=>assessmentDraftIdentity(assessmentFileDraft(),1),()=>assessmentDefinitionFromDraft(assessmentFileDraft(),1)])assert.throws(run,/argument count/);
 for(const key of ['responses','score','consent','reviewed','effectiveTokens']){const f=copy(file);f[key]=[];assert.throws(()=>serializeAssessmentFile(f));const d=copy(file.payload);d[key]=[];assert.throws(()=>createAssessmentFile('definition',d));}
});
