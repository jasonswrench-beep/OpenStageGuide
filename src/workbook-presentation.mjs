// Source-only contracts for a future learner-workbook presentation route.
// Nothing in this module writes Word, changes a course, or renders a DOCX.
import {workbookPreview,WORKBOOK_SCOPE} from './worksheet.mjs';
import {planWorkbookPreview,SELECTED_OUTPUT_SCOPE} from './selected-plan-outputs.mjs';
import {resolveSelectedPlanScope} from './selected-plan-scope.mjs';
export const WORKBOOK_SEMANTIC_VERSION = 1;
export const WORKBOOK_BRAND_POLICY_VERSION = 1;
export const WORKBOOK_MIN_TEXT_CONTRAST = 4.5;
export const WORKBOOK_MAX_SEMANTIC_NODES = 2000;
export const WORKBOOK_MAX_SEMANTIC_TEXT_UNITS = 250000;

const fail = message => { throw new Error(`Unsupported workbook presentation: ${message}`); };
const record = (value, label) => {
  if (!value || typeof value !== 'object' || Array.isArray(value) ||
      ![Object.prototype, null].includes(Object.getPrototypeOf(value))) fail(`${label} must be an own-data record.`);
  if (Object.getOwnPropertySymbols(value).length) fail(`${label} has unsupported symbol fields.`);
  return value;
};
const field = (value, key, label, required = true) => {
  record(value, label);
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  if (!descriptor) {
    if (required) fail(`${label}.${key} is missing.`);
    return undefined;
  }
  if (!Object.hasOwn(descriptor, 'value') || !descriptor.enumerable) fail(`${label}.${key} must be an enumerable data field.`);
  return descriptor.value;
};
const allowed = (value, keys, label) => {
  record(value, label);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key!=='string' || !keys.includes(key)) fail(`${label}.${String(key)} is unsupported.`);
    const descriptor=Object.getOwnPropertyDescriptor(value,key);
    if (!descriptor.enumerable || !Object.hasOwn(descriptor,'value')) fail(`${label}.${key} must be an enumerable data field.`);
  }
};
const list = (value, label) => {
  if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype || value.length > 10000) fail(`${label} must be a dense array.`);
  const keys=Reflect.ownKeys(value);
  if (keys.length!==value.length+1 || !keys.includes('length') ||
      keys.some(key=>key!=='length' && (typeof key!=='string' || !/^(0|[1-9][0-9]*)$/u.test(key) || Number(key)>=value.length))) fail(`${label} must have only dense indexed data.`);
  return Array.from({length: value.length}, (_, i) => {
    const descriptor=Object.getOwnPropertyDescriptor(value,String(i));
    if(!descriptor || !Object.hasOwn(descriptor,'value') || !descriptor.enumerable)fail(`${label}[${i}] must be an enumerable data field.`);
    return descriptor.value;
  });
};
const string = (value, label, max = 50000, required = false) => {
  if (typeof value !== 'string' || value.length > max || (required && !value.trim()) ||
      /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value) || !xmlText(value)) fail(`${label} must be supported text.`);
  return value;
};
const xmlText=value=>{
  for(const symbol of value){
    const cp=symbol.codePointAt(0);
    if(!(cp===9||cp===10||cp===13||cp>=0x20&&cp<=0xd7ff||cp>=0xe000&&cp<=0xfffd||cp>=0x10000&&cp<=0x10ffff))return false;
  }
  return true;
};
const id = (value, label) => string(value, label, 200, true);
const freeze = value => {
  if (value && typeof value === 'object') {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
};

const families = [
  {id:'clearline', name:'Clearline', direction:'spacious-single-column', headingFont:'Aptos Display', bodyFont:'Aptos', accent:'#164B6B'},
  {id:'editorial-study', name:'Editorial Study', direction:'editorial-reference', headingFont:'Georgia', bodyFont:'Aptos', accent:'#543D67'},
  {id:'practice-lab', name:'Practice Lab', direction:'numbered-response', headingFont:'Aptos Display', bodyFont:'Aptos', accent:'#12594D'},
  {id:'field-guide', name:'Field Guide', direction:'procedure-when-semantic', headingFont:'Aptos Display', bodyFont:'Aptos', accent:'#625014'},
  {id:'reflection-journal', name:'Reflection Journal', direction:'prompt-and-response', headingFont:'Georgia', bodyFont:'Aptos', accent:'#5D3D56'}
];
export const WORKBOOK_FAMILIES = freeze(families.map(({id,name,direction,headingFont,bodyFont,accent}) => ({
  id,name,version:3,semanticVersion:WORKBOOK_SEMANTIC_VERSION,direction,accentUse:'decorative-only',
  styles:{title:'Title',module:'Heading1',objective:'Heading2',activity:'Heading2',body:'Normal',response:'Normal'},
  defaults:{organizationName:'',text:'#1A1A1A',background:'#FFFFFF',accent,accentText:'#1A1A1A',accentBackground:'#FFFFFF',headingFont,bodyFont,fallbackFont:'Arial',bodyPt:12,headingPt:16,lineSpacing:1.25,pageSize:'Letter',marginIn:0.8,responseLineIn:0.35,printMode:'color'},
  capabilities:{logo:false,headers:false,footers:false,existingWordTransfer:false,packProfile:false}
})));
export function workbookFamily(familyId) {
  if (typeof familyId !== 'string') fail('family ID must be text.');
  const result = WORKBOOK_FAMILIES.find(f => f.id === familyId);
  if (!result) fail('family ID is unavailable.');
  return result;
}

// Low-level constructor over a public preview and its caller-verified ID map.
// Full-course ID provenance is established by semanticWorkbookFromCourse below.
// Deliberately do not traverse private course fields or reinterpret learner prose.
export function semanticWorkbook(preview, identity, scope) {
  record(preview, 'preview'); record(identity, 'identity'); record(scope, 'scope');
  allowed(identity, ['projectId','revision','modules'], 'identity');
  allowed(scope, ['version','kind','planId','label'], 'scope');
  const version=field(scope,'version','scope',false);
  if(Object.hasOwn(scope,'version') && version!==WORKBOOK_SEMANTIC_VERSION)fail('semantic scope version is unavailable.');
  const kind = field(scope,'kind','scope');
  if (!['full','selected'].includes(kind)) fail('scope kind is unavailable.');
  let textUnits=0, rosterEntries=0;
  const charge=value=>{textUnits+=value.length;if(textUnits>WORKBOOK_MAX_SEMANTIC_TEXT_UNITS)fail('semantic text budget exceeded.');return value;};
  const label = string(field(scope,'label','scope'), 'scope label', 300, true);
  charge(label);
  const projectId = id(field(preview,'projectId','preview'), 'preview project ID');
  charge(projectId);
  if (projectId !== id(field(identity,'projectId','identity'), 'identity project ID')) fail('project identity changed.');
  const revision = field(preview,'revision','preview');
  if (!Number.isSafeInteger(revision) || revision < 0 || revision !== field(identity,'revision','identity')) fail('revision changed.');
  const planId = kind === 'selected' ? id(field(scope,'planId','scope'),'plan ID') : null;
  if(planId)charge(planId);
  if (kind === 'selected' && planId !== id(field(preview,'planId','preview'),'preview plan ID')) fail('plan identity changed.');
  if (kind === 'selected' && label !== string(field(preview,'scopeLabel','preview'),'preview scope label',300,true)) fail('selected scope label changed.');
  if (kind === 'full' && Object.hasOwn(scope,'planId')) fail('full scope cannot carry a plan.');
  const title = string(field(preview,'title','preview'),'course title',50000,true);
  charge(title);
  const groups = list(field(preview,'modules','preview'),'preview modules');
  const mapped = list(field(identity,'modules','identity'),'identity modules');
  if (groups.length !== mapped.length) fail('module identity roster changed.');
  if(kind==='selected'){
    const selectedGroups=list(field(preview,'scope','preview'),'preview selected scope')
      .map(group=>{allowed(group,['moduleId','activityIds'],'preview selected group');
        if(++rosterEntries>WORKBOOK_MAX_SEMANTIC_NODES)fail('selected scope roster budget exceeded.');
        const groupId=charge(id(field(group,'moduleId','preview selected group'),'selected module ID'));
        const activityIds=list(field(group,'activityIds','preview selected group'),'selected activity IDs').map(value=>{
          if(++rosterEntries>WORKBOOK_MAX_SEMANTIC_NODES)fail('selected scope roster budget exceeded.');
          return charge(id(value,'selected activity ID'));
        });
        return {id:groupId,activityIds};})
      .filter(group=>group.activityIds.length);
    if(selectedGroups.length!==mapped.length)fail('selected identity roster changed.');
    for(let i=0;i<mapped.length;i++){
      if(field(mapped[i],'id','identity module')!==selectedGroups[i].id ||
         JSON.stringify(list(field(mapped[i],'activityIds','identity module'),'activity IDs'))!==JSON.stringify(selectedGroups[i].activityIds))fail('selected identity roster changed.');
    }
  }
  const seen = new Set([projectId]);
  const nodes = [];
  const pushId = (value,label) => { const result=id(value,label); if(seen.has(result))fail(`${label} is duplicated.`); seen.add(result); return result; };
  const pushNode=node=>{if(nodes.length>=WORKBOOK_MAX_SEMANTIC_NODES)fail('semantic node budget exceeded.');nodes.push(node);};
  for (let mi=0; mi<groups.length; mi++) {
    const group=groups[mi], map=mapped[mi];
    allowed(group,['title','objective','activities'],'preview module');
    allowed(map,['id','activityIds'],'identity module');
    const moduleId=pushId(field(map,'id','identity module'),'module ID');
    const moduleTitle=charge(string(field(group,'title','preview module'),'module title',50000,true));
    const objective=charge(string(field(group,'objective','preview module'),'module objective'));
    charge(moduleId);pushNode({kind:'module',id:moduleId,title:moduleTitle,objective});
    const activities=list(field(group,'activities','preview module'),'preview activities');
    const activityIds=list(field(map,'activityIds','identity module'),'activity IDs');
    if(activities.length!==activityIds.length)fail('activity identity roster changed.');
    for(let ai=0;ai<activities.length;ai++){
      const activity=activities[ai];allowed(activity,['title','participant','lines'],'preview activity');
      const lines=field(activity,'lines','preview activity');
      if(!Number.isInteger(lines)||lines<0||lines>20)fail('writing lines must be 0–20.');
      const activityId=pushId(activityIds[ai],'activity ID');
      const activityTitle=charge(string(field(activity,'title','preview activity'),'activity title',50000,true));
      const instructions=charge(string(field(activity,'participant','preview activity'),'participant instructions'));
      charge(activityId);pushNode({kind:'activity',id:activityId,moduleId,title:activityTitle,instructions,responseLines:lines});
    }
  }
  const identityLines=kind==='selected' ? list(field(preview,'identityLines','preview'),'identity lines').map(x=>charge(string(x,'identity line',50000,true))) : [];
  return freeze({version:WORKBOOK_SEMANTIC_VERSION,projectId,revision,scope:{kind,planId,label},title,identityLines,nodes});
}

// Derive both learner text and ordered IDs synchronously from the same authoring
// project. These pure adapters do not bind a future asynchronous UI review.
export function semanticWorkbookFromCourse(project){
  const preview=workbookPreview(project);
  const modules=list(field(project,'modules','project'),'project modules').map(module=>({
    id:field(module,'id','project module'),
    activityIds:list(field(module,'activities','project module'),'project activities').map(activity=>field(activity,'id','project activity'))
  }));
  return semanticWorkbook(preview,{projectId:field(project,'id','project'),revision:field(project,'revision','project'),modules},{kind:'full',label:WORKBOOK_SCOPE});
}
export function semanticWorkbookFromSelected(project,planId,stamp,start='09:00'){
  const preview=planWorkbookPreview(project,planId,stamp,start);
  const resolved=resolveSelectedPlanScope(project,planId,start);
  const modules=resolved.groups.map(group=>({id:group.moduleId,activityIds:[...group.activityIds]}));
  return semanticWorkbook(preview,{projectId:resolved.projectId,revision:resolved.revision,modules},{kind:'selected',planId:resolved.planId,label:SELECTED_OUTPUT_SCOPE});
}

const fonts=new Set(['Aptos','Aptos Display','Arial','Calibri','Georgia','Segoe UI','Times New Roman']);
const brandKeys=['organizationName','text','background','accent','accentText','accentBackground','headingFont','bodyFont','fallbackFont','bodyPt','headingPt','lineSpacing','pageSize','marginIn','responseLineIn','printMode'];
const hex=value=>typeof value==='string'&&/^#[0-9a-fA-F]{6}$/u.test(value);
const luminance=color=>{
  const rgb=[1,3,5].map(i=>parseInt(color.slice(i,i+2),16)/255).map(v=>v<=0.04045?v/12.92:((v+0.055)/1.055)**2.4);
  return 0.2126*rgb[0]+0.7152*rgb[1]+0.0722*rgb[2];
};
export function textContrast(foreground,background){
  if(!hex(foreground)||!hex(background))fail('contrast needs opaque six-digit colors.');
  const a=luminance(foreground),b=luminance(background);
  return (Math.max(a,b)+0.05)/(Math.min(a,b)+0.05);
}
const safeText=background=>textContrast('#000000',background)>=WORKBOOK_MIN_TEXT_CONTRAST?'#000000':'#FFFFFF';
export function resolveWorkbookBrand(familyId, request = {}) {
  const family=workbookFamily(familyId);
  allowed(request,brandKeys,'brand request');
  const rawRequest={};
  for(const key of Reflect.ownKeys(request)) {
    const value=field(request,key,'brand request');
    if(value!==null && !['string','number','boolean'].includes(typeof value))fail(`brand request.${key} must be a JSON scalar value.`);
    if(typeof value==='number'&&!Number.isFinite(value))fail(`brand request.${key} must be finite.`);
    rawRequest[key]=value;
  }
  const issues=[];
  const proposed={};
  const effective={...family.defaults,...rawRequest};
  for(const key of ['text','background','accent','accentText','accentBackground'])
    if(!hex(effective[key]))issues.push(`${key} must be an opaque #RRGGBB color.`);
  if(hex(effective.text)&&hex(effective.background)&&textContrast(effective.text,effective.background)<WORKBOOK_MIN_TEXT_CONTRAST){
    issues.push('Body text/background contrast is below 4.5:1.');proposed.text=safeText(effective.background);
  }
  if(hex(effective.accentText)&&hex(effective.accentBackground)&&textContrast(effective.accentText,effective.accentBackground)<WORKBOOK_MIN_TEXT_CONTRAST){
    issues.push('Accent text/background contrast is below 4.5:1.');proposed.accentText=safeText(effective.accentBackground);
  }
  if(typeof effective.organizationName!=='string'||effective.organizationName.length>120||!xmlText(effective.organizationName))issues.push('Organization name must be XML-safe text of at most 120 characters.');
  for(const key of ['headingFont','bodyFont','fallbackFont'])if(!fonts.has(effective[key]))issues.push(`${key} must be an approved font name; installed fonts are not verified here.`);
  for(const [key,min,max] of [['bodyPt',11,18],['headingPt',14,28],['lineSpacing',1,2],['marginIn',0.5,1.5],['responseLineIn',0.25,0.7]])
    if(typeof effective[key]!=='number'||!Number.isFinite(effective[key])||effective[key]<min||effective[key]>max)issues.push(`${key} is outside the supported range.`);
  if(!['Letter','A4'].includes(effective.pageSize))issues.push('Page size must be Letter or A4.');
  if(!['color','grayscale'].includes(effective.printMode))issues.push('Print mode is unsupported.');
  return freeze({eligible:issues.length===0,policyVersion:WORKBOOK_BRAND_POLICY_VERSION,familyId:family.id,familyVersion:family.version,
    rawRequest, effective:issues.length?null:effective, issues, proposedSafeAlternative:proposed});
}
