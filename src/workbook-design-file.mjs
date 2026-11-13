import {FAMILY_BRAND_FIELDS} from './workbook-family-flow.mjs';
import {WORKBOOK_FAMILIES} from './workbook-presentation.mjs';
import {WORKBOOK_BASE_LANGUAGES} from './workbook-family-output.mjs';
import {WORKBOOK_COVER_FIELDS,WORKBOOK_COPYRIGHT_FIELDS,WORKBOOK_NOTICE_CHOICES} from './workbook-public-frontmatter.mjs';
import {WORKBOOK_RESEARCH_FIELDS} from './workbook-research.mjs';
import {WORKBOOK_LOGO_VERSION,WORKBOOK_LOGO_LIMITS,prepareWorkbookLogo,workbookLogoInfo,workbookLogoBytes,workbookLogoBase64} from './workbook-logo.mjs';

// Explicit PRIVATE raw-design recovery files. No course, assessment, storage,
// network, consent, output review or document-write capability is supplied.
export const WORKBOOK_DESIGN_FILE_FORMAT='openstage-workbook-design';
export const WORKBOOK_DESIGN_FILE_VERSION=1;
export const WORKBOOK_DESIGN_FILE_LIMITS=Object.freeze({maxFileBytes:6*1024*1024,maxRawTextUnits:200000,maxJsonDepth:16,maxJsonNodes:100000,maxPngBytes:WORKBOOK_LOGO_LIMITS.bytes,maxBase64Units:4*Math.ceil(WORKBOOK_LOGO_LIMITS.bytes/3)});
const TOP=['familyId','baseLanguage','brand','frontMatter','contents','research','logo'];
const LOGO=['include','role','alt','widthIn'];
const FONTS=['','Aptos','Aptos Display','Arial','Calibri','Georgia','Segoe UI','Times New Roman'];
const brandCaps={organizationName:120,text:7,background:7,accent:7,accentText:7,accentBackground:7,headingFont:80,bodyFont:80,fallbackFont:80,bodyPt:12,headingPt:12,lineSpacing:12,pageSize:20,marginIn:12,responseLineIn:12,printMode:20};
const coverCaps={title:5000,subtitle:8000,author:1000,edition:500,year:32,audience:5000};
const copyrightCaps={rightsHolder:1000,year:32,noticeChoice:32,noticeText:16000,permissionsContact:2000,credits:16000,licenseName:200,licenseVersion:80,licenseUrl:2048};
const fail=m=>{throw new Error(`Unsupported PRIVATE workbook-design file: ${m}`);};
const arity=(n,wanted)=>{if(n!==wanted)fail('argument count changed.');};
function record(value,keys,label){
 if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be plain own data.`);
 const found=Reflect.ownKeys(value),out=Object.create(null);
 if(found.length!==keys.length||found.some(k=>typeof k!=='string'||!keys.includes(k)))fail(`${label} fields changed.`);
 for(const key of keys){const d=Object.getOwnPropertyDescriptor(value,key);if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${key} must be visible own data.`);out[key]=d.value;}
 return out;
}
function freeze(v){if(v&&typeof v==='object'){for(const child of Object.values(v))freeze(child);Object.freeze(v);}return v;}
function scalar(v,label,b,max){
 if(typeof v!=='string'||v.length>max)fail(`${label} must be bounded raw text.`);
 for(let i=0;i<v.length;i++){const c=v.charCodeAt(i);if(c>=0xd800&&c<=0xdbff){const n=v.charCodeAt(++i);if(!(n>=0xdc00&&n<=0xdfff))fail(`${label} has an unpaired surrogate.`);}else if(c>=0xdc00&&c<=0xdfff)fail(`${label} has an unpaired surrogate.`);}
 b.units+=v.length;if(b.units>WORKBOOK_DESIGN_FILE_LIMITS.maxRawTextUnits)fail('raw text budget exceeded.');return v;
}
function bool(v,label){if(typeof v!=='boolean')fail(`${label} must be boolean data.`);return v;}
function choice(v,values,label){if(!values.includes(v))fail(`${label} choice is unavailable.`);return v;}
function fields(v,keys,caps,b,label){const r=record(v,keys,label);return Object.fromEntries(keys.map(key=>[key,scalar(r[key],`${label}.${key}`,b,caps[key]) ]));}
function normalizeDraft(value){
 const r=record(value,TOP,'draft'),b={units:0};
 const familyId=choice(scalar(r.familyId,'family ID',b,80),WORKBOOK_FAMILIES.map(f=>f.id),'family');
 const baseLanguage=choice(scalar(r.baseLanguage,'base language',b,16),WORKBOOK_BASE_LANGUAGES,'base language');
 const brand=fields(r.brand,FAMILY_BRAND_FIELDS,brandCaps,b,'brand');
 for(const key of ['headingFont','bodyFont','fallbackFont'])choice(brand[key],FONTS,key);
 choice(brand.pageSize,['','Letter','A4'],'page size');choice(brand.printMode,['','color','grayscale'],'print mode');
 const f=record(r.frontMatter,['includeCover','includeCopyright','cover','copyright'],'front matter');
 const frontMatter={includeCover:bool(f.includeCover,'cover inclusion'),includeCopyright:bool(f.includeCopyright,'copyright inclusion'),cover:fields(f.cover,WORKBOOK_COVER_FIELDS,coverCaps,b,'cover'),copyright:fields(f.copyright,WORKBOOK_COPYRIGHT_FIELDS,copyrightCaps,b,'copyright')};
 choice(frontMatter.copyright.noticeChoice,['',...WORKBOOK_NOTICE_CHOICES],'notice');
 const c=record(r.contents,['include'],'contents'),contents={include:bool(c.include,'contents inclusion')};
 const researchSource=record(r.research,['include',...WORKBOOK_RESEARCH_FIELDS],'research');
 const research={include:bool(researchSource.include,'research inclusion'),...Object.fromEntries(WORKBOOK_RESEARCH_FIELDS.map(key=>[key,scalar(researchSource[key],`research.${key}`,b,16000)]))};
 const l=record(r.logo,LOGO,'logo'),logo={include:bool(l.include,'logo inclusion'),role:choice(scalar(l.role,'logo role',b,20),['','informative','decorative'],'logo role'),alt:scalar(l.alt,'logo alternative description',b,WORKBOOK_LOGO_LIMITS.altTextUnits),widthIn:scalar(l.widthIn,'raw logo width',b,32)};
 return {familyId,baseLanguage,brand,frontMatter,contents,research,logo};
}
/** Raw drafts remain editable; validation is not workbook export eligibility. */
export function validateWorkbookDesignDraft(value){arity(arguments.length,1);return freeze(normalizeDraft(value));}
const alphabet='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
function image(value){
 if(value===null)return null;
 const r=record(value,['mediaType','normalizerVersion','sha256','base64'],'PNG image');
 if(r.mediaType!=='image/png'||r.normalizerVersion!==WORKBOOK_LOGO_VERSION||typeof r.sha256!=='string'||!/^[0-9a-f]{64}$/u.test(r.sha256))fail('PNG image type, version or integrity metadata changed.');
 const s=r.base64;
 if(typeof s!=='string'||!s.length||s.length>WORKBOOK_DESIGN_FILE_LIMITS.maxBase64Units||s.length%4)fail('PNG base64 is unsupported or too large.');
 const padding=s.endsWith('==')?2:s.endsWith('=')?1:0,bytes=s.length/4*3-padding;
 for(let i=0;i<s.length-padding;i++){const c=s.charCodeAt(i);if(!(c>=65&&c<=90||c>=97&&c<=122||c>=48&&c<=57||c===43||c===47))fail('PNG base64 contains unsupported characters or padding.');}
 if(!bytes||bytes>WORKBOOK_LOGO_LIMITS.bytes||padding===2&&(alphabet.indexOf(s.at(-3))&15)||padding===1&&(alphabet.indexOf(s.at(-2))&3))fail('PNG base64 is noncanonical or exceeds 2 MiB.');
 return {mediaType:'image/png',normalizerVersion:WORKBOOK_LOGO_VERSION,sha256:r.sha256,base64:s};
}
function normalizeFile(value){
 const f=record(value,['format','version','payload'],'file');if(f.format!==WORKBOOK_DESIGN_FILE_FORMAT||f.version!==WORKBOOK_DESIGN_FILE_VERSION)fail('file format/version is unavailable.');
 const p=record(f.payload,TOP,'payload'),l=record(p.logo,[...LOGO,'image'],'file logo');
 const raw=normalizeDraft({...p,logo:Object.fromEntries(LOGO.map(key=>[key,l[key]]))});
 return {format:WORKBOOK_DESIGN_FILE_FORMAT,version:WORKBOOK_DESIGN_FILE_VERSION,payload:{...raw,logo:{...raw.logo,image:image(l.image)}}};
}
/** Save source is an owned verified current handle or explicit absent image. */
export function createWorkbookDesignFile(value,logoAsset){
 arity(arguments.length,2);const draft=normalizeDraft(value);let captured=null;
 if(logoAsset!==null){const info=workbookLogoInfo(logoAsset);captured={mediaType:'image/png',normalizerVersion:WORKBOOK_LOGO_VERSION,sha256:info.sha256,base64:workbookLogoBase64(logoAsset)};}
 return freeze(normalizeFile({format:WORKBOOK_DESIGN_FILE_FORMAT,version:WORKBOOK_DESIGN_FILE_VERSION,payload:{...draft,logo:{...draft.logo,image:captured}}}));
}
/** Canonical UTF-8 authoring JSON; this does not adopt or approve a PNG. */
export function serializeWorkbookDesignFile(value){
 arity(arguments.length,1);const f=normalizeFile(value),bytes=new TextEncoder().encode(JSON.stringify(f,null,2)+'\n');if(bytes.length>WORKBOOK_DESIGN_FILE_LIMITS.maxFileBytes)fail('file exceeds 6 MiB.');return bytes;
}
function strictJson(text){
 let i=0,nodes=0;const ws=()=>{while(i<text.length&&/[\x20\t\r\n]/u.test(text[i]))i++;};
 function string(key=false){const start=i;if(text[i++]!=='"')fail('invalid JSON string.');while(i<text.length){const c=text.charCodeAt(i++);if(c===34){try{return key?JSON.parse(text.slice(start,i)):undefined;}catch{fail('invalid JSON string.');}}if(c<32)fail('unescaped JSON control character.');if(c===92){const e=text[i++];if(e==='u'){if(!/^[0-9a-fA-F]{4}$/u.test(text.slice(i,i+4)))fail('invalid JSON escape.');i+=4;}else if(!['"','\\','/','b','f','n','r','t'].includes(e))fail('invalid JSON escape.');}}fail('unterminated JSON string.');}
 function value(depth){if(++nodes>WORKBOOK_DESIGN_FILE_LIMITS.maxJsonNodes)fail('JSON node budget exceeded.');if(depth>WORKBOOK_DESIGN_FILE_LIMITS.maxJsonDepth)fail('JSON nesting exceeds this file profile.');ws();const c=text[i];if(c==='{'){i++;ws();const keys=new Set();if(text[i]==='}'){i++;return;}for(;;){ws();const key=string(true);if(keys.has(key))fail('duplicate JSON object key.');keys.add(key);ws();if(text[i++]!==':')fail('invalid JSON object.');value(depth+1);ws();const end=text[i++];if(end==='}')return;if(end!==',')fail('invalid JSON object.');}}else if(c==='['){i++;ws();if(text[i]===']'){i++;return;}for(;;){value(depth+1);ws();const end=text[i++];if(end===']')return;if(end!==',')fail('invalid JSON array.');}}else if(c==='"')string();else{const start=i;while(i<text.length&&!/[\x20\t\r\n,\]}]/u.test(text[i]))i++;if(!/^(?:true|false|null|-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?)$/u.test(text.slice(start,i)))fail('invalid JSON value.');}}
 value(0);ws();if(i!==text.length)fail('trailing JSON data.');try{return JSON.parse(text);}catch{fail('invalid JSON.');}
}
const typedLength=Object.getOwnPropertyDescriptor(Object.getPrototypeOf(Uint8Array.prototype),'byteLength').get;
export function parseWorkbookDesignFile(input){
 arity(arguments.length,1);if(!(input instanceof Uint8Array))fail('provide UTF-8 file bytes.');let length;try{length=typedLength.call(input);}catch{fail('file bytes are unavailable.');}if(!length||length>WORKBOOK_DESIGN_FILE_LIMITS.maxFileBytes)fail('choose a nonempty workbook-design file up to 6 MiB.');
 const bytes=new Uint8Array(input);if(bytes[0]===0xef&&bytes[1]===0xbb&&bytes[2]===0xbf)fail('UTF-8 BOM is outside this file profile.');let text;try{text=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch{fail('file is not valid UTF-8.');}return freeze(normalizeFile(strictJson(text)));
}
function decodeImage(source){
 if(typeof globalThis.atob!=='function')fail('PNG base64 decoding is unavailable.');let binary;try{binary=globalThis.atob(source.base64);}catch{fail('PNG base64 decoding failed.');}
 if(binary.length>WORKBOOK_LOGO_LIMITS.bytes)fail('decoded PNG exceeds 2 MiB.');const out=new Uint8Array(binary.length);for(let i=0;i<binary.length;i++)out[i]=binary.charCodeAt(i);return out;
}
/** Verify stored integrity, then re-normalize into a new owned session handle. */
export async function prepareWorkbookDesignFile(value){
 arity(arguments.length,1);const file=freeze(normalizeFile(value)),canonical=JSON.stringify(file),p=file.payload,{image:stored,...rawLogo}=p.logo;
 const draft=freeze({...p,logo:rawLogo});let logoAsset=null,logoInfo=null,normalization=null;
 if(stored){
  const bytes=decodeImage(stored);if(typeof globalThis.crypto?.subtle?.digest!=='function')fail('PNG integrity checking is unavailable.');
  const digest=new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256',new Uint8Array(bytes))),sha256=Array.from(digest,v=>v.toString(16).padStart(2,'0')).join('');
  if(sha256!==stored.sha256)fail('stored PNG integrity check failed.');
  logoAsset=await prepareWorkbookLogo(new Uint8Array(bytes));logoInfo=workbookLogoInfo(logoAsset);const current=workbookLogoBytes(logoAsset);
  normalization={storedSha256:stored.sha256,currentSha256:logoInfo.sha256,bytesChanged:bytes.length!==current.length||bytes.some((v,i)=>v!==current[i])};
 }
 if(JSON.stringify(normalizeFile(value))!==canonical)fail('file source changed during PNG validation. Nothing was prepared.');
 return freeze({file,draft,logoAsset,logoInfo,normalization});
}
