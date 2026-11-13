import {FAMILY_BRAND_FIELDS} from './workbook-family-flow.mjs';

export const BRAND_PRESET_KEY='openstage-guide.brand-presets.v1';
export const BRAND_PRESET_VERSION=1;
export const MAX_BRAND_PRESETS=20;
export const MAX_BRAND_PRESET_BYTES=32768;
const empty=()=>({format:'openstage-guide-brand-presets',version:BRAND_PRESET_VERSION,revision:'empty',presets:[]});
const fail=message=>{throw new Error(`Brand presets: ${message}`);};
const own=(value,keys,label)=>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be a plain record.`);
  const found=Reflect.ownKeys(value);
  if(found.length!==keys.length||found.some(key=>typeof key!=='string'||!keys.includes(key)))fail(`${label} fields are unsupported.`);
  for(const key of keys){const d=Object.getOwnPropertyDescriptor(value,key);
    if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${key} must be visible own data.`);
  }
  return value;
};
const idPattern=/^[0-9a-f]{8}-[0-9a-f]{4}-[1-8][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/u;
const dense=(value,label,max)=>{
  if(!Array.isArray(value)||Object.getPrototypeOf(value)!==Array.prototype||value.length>max)fail(`${label} must be a bounded dense list.`);
  const keys=Reflect.ownKeys(value);
  if(keys.length!==value.length+1||!keys.includes('length')||keys.some(key=>key!=='length'&&
      (typeof key!=='string'||!/^(0|[1-9][0-9]*)$/u.test(key)||Number(key)>=value.length)))fail(`${label} has unsupported fields.`);
  for(let i=0;i<value.length;i++){
    const d=Object.getOwnPropertyDescriptor(value,String(i));
    if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}[${i}] must be own data.`);
  }
  return value;
};
export function validatePresetBrand(brand){
  own(brand,FAMILY_BRAND_FIELDS,'brand');
  const result={};
  for(const key of FAMILY_BRAND_FIELDS){
    const value=brand[key];
    if(typeof value!=='string'||value.length>(key==='organizationName'?120:80)||/[\u0000-\u001f\u007f]/u.test(value))fail(`${key} is not a bounded raw override.`);
    result[key]=value;
  }
  return result;
}
export function validateBrandPresetCatalog(value){
  own(value,['format','version','revision','presets'],'catalog');
  if(value.format!=='openstage-guide-brand-presets'||value.version!==BRAND_PRESET_VERSION)fail('the saved catalog version is unsupported; it was not changed.');
  if(typeof value.revision!=='string'||value.revision!=='empty'&&!idPattern.test(value.revision))fail('catalog revision is invalid.');
  dense(value.presets,'presets',MAX_BRAND_PRESETS);
  const ids=new Set(),names=new Set(),presets=[];
  for(let i=0;i<value.presets.length;i++){
    const row=value.presets[i];
    own(row,['id','name','brand'],'preset');
    if(typeof row.id!=='string'||!idPattern.test(row.id)||ids.has(row.id))fail('preset ID is missing or duplicated.');
    if(typeof row.name!=='string'||row.name!==row.name.trim()||!row.name||row.name.length>80||/[\u0000-\u001f\u007f]/u.test(row.name))fail('preset name is invalid.');
    const folded=row.name.toLocaleLowerCase('en-US');
    if(names.has(folded))fail('preset name is duplicated.');
    ids.add(row.id);names.add(folded);
    presets.push({id:row.id,name:row.name,brand:validatePresetBrand(row.brand)});
  }
  return {format:value.format,version:value.version,revision:value.revision,presets};
}
export function readBrandPresetCatalog(storage){
  if(!storage||typeof storage.getItem!=='function')fail('local storage is unavailable; current branding was kept.');
  let raw;
  try{raw=storage.getItem(BRAND_PRESET_KEY);}catch{fail('local storage could not be read; current branding was kept.');}
  if(raw===null)return {raw:null,catalog:empty()};
  if(typeof raw!=='string'||new TextEncoder().encode(raw).length>MAX_BRAND_PRESET_BYTES)fail('saved catalog is oversized or invalid; it was not changed.');
  let parsed;
  try{parsed=JSON.parse(raw);}catch{fail('saved catalog is malformed; it was not changed.');}
  const catalog=validateBrandPresetCatalog(parsed);
  if(JSON.stringify(catalog)!==raw)fail('saved catalog is noncanonical or altered; it was not changed.');
  return {raw,catalog};
}
export function writeBrandPresetCatalog(storage,baselineRaw,next){
  if(typeof baselineRaw!=='string'&&baselineRaw!==null)fail('catalog baseline is invalid.');
  const current=readBrandPresetCatalog(storage);
  if(current.raw!==baselineRaw)fail('the saved presets changed in another session; reload them before writing.');
  const catalog=validateBrandPresetCatalog(next),raw=JSON.stringify(catalog);
  if(new TextEncoder().encode(raw).length>MAX_BRAND_PRESET_BYTES)fail('catalog exceeds the local storage limit.');
  try{storage.setItem(BRAND_PRESET_KEY,raw);}catch{fail('write outcome is unconfirmed; current branding and last confirmed preset list were kept. Reload saved presets before another write.');}
  let readback;
  try{readback=storage.getItem(BRAND_PRESET_KEY);}catch{fail('write outcome is unconfirmed; current branding and last confirmed preset list were kept. Reload saved presets before another write.');}
  if(readback!==raw)fail('write outcome is unconfirmed; current branding and last confirmed preset list were kept. Reload saved presets before another write.');
  return {raw,catalog};
}
export function nextBrandPresetCatalog(catalog,action,options){
  own(options,['id','name','brand','uuid'],'preset operation');
  const {id,name,brand,uuid}=options;
  const old=validateBrandPresetCatalog(catalog);
  if(typeof uuid!=='string'||!idPattern.test(uuid))fail('a fresh catalog revision is required.');
  const presets=old.presets.map(row=>({id:row.id,name:row.name,brand:{...row.brand}}));
  if(action==='create'){
    if(typeof id!=='string'||!idPattern.test(id))fail('a fresh preset ID is required.');
    if(presets.length>=MAX_BRAND_PRESETS)fail('the preset limit is reached.');
    presets.push({id,name,brand});
  }else if(action==='overwrite'){
    const index=presets.findIndex(row=>row.id===id);
    if(index<0)fail('the selected preset no longer exists.');
    presets[index]={id,name,brand};
  }else fail('this preset action is unsupported.');
  return validateBrandPresetCatalog({format:old.format,version:old.version,revision:uuid,presets});
}
