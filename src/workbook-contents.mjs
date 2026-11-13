// Generated-copy internal navigation only. No course, Word, storage, service or
// page-layout calculation is read or performed by this pure module.
export const WORKBOOK_CONTENTS_VERSION=1;
export const WORKBOOK_CONTENTS_NOTE='Internal navigation links only. Page numbers are not calculated in this generated copy. This snapshot does not automatically update entry text after manual edits and is not an updating Word contents field. Regenerate from saved source after headings change.';
export const WORKBOOK_CONTENTS_VALIDITY=Object.freeze(['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing']);
export const WORKBOOK_CONTENTS_LIMITS=Object.freeze({headings:4007,labelTextUnits:400000,identityTextUnits:1000000,keyTextUnits:1000,labelFieldTextUnits:50000});
const MAX_HEADINGS=WORKBOOK_CONTENTS_LIMITS.headings;
const fail=message=>{throw new Error(`Workbook contents: ${message}`);};
const own=(value,keys,label)=>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be an own-data record.`);
  const found=Reflect.ownKeys(value);
  if(found.length!==keys.length||found.some(k=>typeof k!=='string'||!keys.includes(k)))fail(`${label} fields are unsupported.`);
  for(const k of keys){const d=Object.getOwnPropertyDescriptor(value,k);if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${k} must be enumerable data.`);}
};
const text=(value,max,label)=>{
  if(typeof value!=='string'||!value.trim()||value.length>max)fail(`${label} must be bounded nonblank text.`);
  for(const symbol of value){const cp=symbol.codePointAt(0);if(!(cp===9||cp===10||cp===13||cp>=0x20&&cp<=0xd7ff||cp>=0xe000&&cp<=0xfffd||cp>=0x10000&&cp<=0x10ffff))fail(`${label} contains unsupported XML text.`);}
  return value.replace(/\r\n?/gu,'\n');
};
const freeze=value=>{if(value&&typeof value==='object'){for(const v of Object.values(value))freeze(v);Object.freeze(value);}return value;};
export function validateWorkbookContents(value){
  if(arguments.length!==1)fail('provide one contents configuration.');
  own(value,['schemaVersion','include'],'configuration');
  if(value.schemaVersion!==1||typeof value.include!=='boolean')fail('version or inclusion is unsupported.');
  return Object.freeze({schemaVersion:1,include:value.include});
}
export function workbookContentsRequest(raw){
  if(arguments.length!==1)fail('provide one raw inclusion control.');
  own(raw,['checked','value','invalid'],'inclusion control');own(raw.invalid,WORKBOOK_CONTENTS_VALIDITY,'inclusion validity');
  if(typeof raw.checked!=='boolean'||typeof raw.value!=='string'||raw.value.length>100||WORKBOOK_CONTENTS_VALIDITY.some(k=>typeof raw.invalid[k]!=='boolean'||raw.invalid[k]))fail('the contents inclusion control is invalid; its draft was retained.');
  return validateWorkbookContents({schemaVersion:1,include:raw.checked});
}
// Two differently seeded 64-bit FNV streams, truncated to96bits, keep Word names
// short. Collisions are explicitly grouped and key-sorted below; the digest is
// neither a signature nor a claim of global cross-document identity.
const digest=value=>{
  let a=0xcbf29ce484222325n,b=0x84222325cbf29ce4n;
  for(const byte of new TextEncoder().encode(value)){a=BigInt.asUintN(64,(a^BigInt(byte))*0x100000001b3n);b=BigInt.asUintN(64,(b^BigInt(byte))*0x100000001b3n);}
  return (a.toString(16).padStart(16,'0')+b.toString(16).padStart(16,'0')).slice(0,24);
};
export function buildWorkbookContents(headings){
  if(arguments.length!==1)fail('provide one emitted heading roster.');
  if(!Array.isArray(headings)||Object.getPrototypeOf(headings)!==Array.prototype||headings.length>MAX_HEADINGS)fail('heading roster is unsupported.');
  const keys=Reflect.ownKeys(headings);if(keys.length!==headings.length+1||keys.some(k=>k!=='length'&&(typeof k!=='string'||!/^(0|[1-9][0-9]*)$/u.test(k)||Number(k)>=headings.length)))fail('heading roster must be dense own data.');
  const seen=new Set(),rows=[];let labelUnits=0,keyUnits=0;
  for(let i=0;i<headings.length;i++){
    const d=Object.getOwnPropertyDescriptor(headings,String(i));if(!d?.enumerable||!Object.hasOwn(d,'value'))fail('heading roster must be dense own data.');
    own(d.value,['key','level','text'],'heading');const row=d.value,key=text(row.key,WORKBOOK_CONTENTS_LIMITS.keyTextUnits,'heading identity'),label=text(row.text,WORKBOOK_CONTENTS_LIMITS.labelFieldTextUnits,'heading label');
    if(seen.has(key)||![1,2].includes(row.level))fail('heading identity is duplicated or level unsupported.');seen.add(key);labelUnits+=label.length;keyUnits+=key.length;if(labelUnits>WORKBOOK_CONTENTS_LIMITS.labelTextUnits||keyUnits>WORKBOOK_CONTENTS_LIMITS.identityTextUnits)fail('heading roster text budget exceeded.');
    rows.push({key,level:row.level,text:label,base:'OSG_'+digest(key)});
  }
  const collisions=new Map();for(const row of rows){if(!collisions.has(row.base))collisions.set(row.base,[]);collisions.get(row.base).push(row.key);}
  for(const group of collisions.values())group.sort();
  const names=new Set(),entries=rows.map((row,i)=>{
    const group=collisions.get(row.base),anchor=row.base+(group.length>1?'_'+(group.indexOf(row.key)+1):'');
    if(anchor.length>40||!/^[A-Za-z][A-Za-z0-9_]*$/u.test(anchor)||names.has(anchor))fail('bookmark naming is unsupported.');names.add(anchor);
    return {level:row.level,text:row.text,anchor,bookmarkId:i+1};
  });
  return freeze({schemaVersion:1,title:'Contents',note:WORKBOOK_CONTENTS_NOTE,pageNumbers:false,entries});
}
