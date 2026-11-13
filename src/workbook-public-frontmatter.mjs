
// Public, generated-copy metadata only. This module never reads a course or
// facilitator record and does not persist a draft or choose a legal license.
export const WORKBOOK_PUBLIC_FRONT_MATTER_VERSION=1;
export const WORKBOOK_COVER_FIELDS=Object.freeze(['title','subtitle','author','edition','year','audience']);
export const WORKBOOK_COPYRIGHT_FIELDS=Object.freeze(['rightsHolder','year','noticeChoice','noticeText','permissionsContact','credits','licenseName','licenseVersion','licenseUrl']);
export const WORKBOOK_NOTICE_CHOICES=Object.freeze(['all-rights-reserved','custom-permissions','named-license','custom-text']);
export const WORKBOOK_FRONT_VALIDITY=Object.freeze(['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing']);
const limits=Object.freeze({title:5000,subtitle:8000,author:1000,edition:500,year:32,audience:5000,rightsHolder:1000,noticeChoice:32,noticeText:16000,permissionsContact:2000,credits:16000,licenseName:200,licenseVersion:80,licenseUrl:2048});
const freeze=value=>{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value);}return value;};
export const EMPTY_WORKBOOK_FRONT_MATTER=freeze({schemaVersion:1,includeCover:false,includeCopyright:false,
  cover:Object.fromEntries(WORKBOOK_COVER_FIELDS.map(k=>[k,''])),copyright:Object.fromEntries(WORKBOOK_COPYRIGHT_FIELDS.map(k=>[k,'']))});
const fail=message=>{throw new Error(`Public workbook front matter: ${message}`);};
const own=(value,keys,label)=>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be an own-data record.`);
  const names=Reflect.ownKeys(value);
  if(names.length!==keys.length||names.some(k=>typeof k!=='string'||!keys.includes(k)))fail(`${label} contains unsupported fields.`);
  for(const key of keys){const d=Object.getOwnPropertyDescriptor(value,key);if(!d?.enumerable||!Object.hasOwn(d,'value'))fail(`${label}.${key} must be enumerable data.`);}
  return value;
};
const supportedText=(value,label,max)=>{
  if(typeof value!=='string'||value.length>max)fail(`${label} must be bounded text.`);
  for(const symbol of value){const cp=symbol.codePointAt(0);if(!(cp===9||cp===10||cp===13||cp>=0x20&&cp<=0xd7ff||cp>=0xe000&&cp<=0xfffd||cp>=0x10000&&cp<=0x10ffff))fail(`${label} contains unsupported XML text.`);}
  return value.replace(/\r\n?/gu,'\n');
};
const text=(value,key)=>supportedText(value,key,limits[key]);
const year=(value,label,required)=>{
  if(!value){if(required)fail(`${label} is required when its page is included.`);return;}
  const match=/^\s*([0-9]{4})(?:\s*[-\u2013]\s*([0-9]{4}))?\s*$/u.exec(value);
  if(!match||Number(match[1])<1000||match[2]&&(Number(match[2])<1000||Number(match[2])<Number(match[1])))fail(`${label} must be a four-digit year or an ordered year range.`);
};
const licenseUrl=value=>{
  if(!value||value!==value.trim()||/\s/u.test(value))fail('named license URL must be an explicit HTTP or HTTPS URL without credentials.');
  let url;try{url=new URL(value);}catch{fail('named license URL must be an explicit HTTP or HTTPS URL without credentials.');}
  if(!['http:','https:'].includes(url.protocol)||url.username||url.password)fail('named license URL must be an explicit HTTP or HTTPS URL without credentials.');
};
export function validateWorkbookFrontMatter(value){
  if(arguments.length!==1)fail('provide one metadata record only.');
  own(value,['schemaVersion','includeCover','includeCopyright','cover','copyright'],'metadata');
  if(value.schemaVersion!==1||typeof value.includeCover!=='boolean'||typeof value.includeCopyright!=='boolean')fail('metadata version and inclusion choices are unsupported.');
  own(value.cover,WORKBOOK_COVER_FIELDS,'cover');own(value.copyright,WORKBOOK_COPYRIGHT_FIELDS,'copyright');
  const cover=Object.fromEntries(WORKBOOK_COVER_FIELDS.map(k=>[k,text(value.cover[k],k)]));
  const copyright=Object.fromEntries(WORKBOOK_COPYRIGHT_FIELDS.map(k=>[k,text(value.copyright[k],k)]));
  if(Object.values(cover).concat(Object.values(copyright)).reduce((n,v)=>n+v.length,0)>64000)fail('metadata exceeds the 64,000 text-unit limit.');
  if(copyright.noticeChoice!==''&&!WORKBOOK_NOTICE_CHOICES.includes(copyright.noticeChoice))fail('notice choice is unsupported.');
  if(value.includeCover){if(!cover.title.trim())fail('enter an explicit public cover title before including the cover.');year(cover.year,'cover year',false);}
  if(value.includeCopyright){
    if(!copyright.rightsHolder.trim())fail('enter the rights holder separately from the cover author.');
    year(copyright.year,'copyright year',true);
    if(!WORKBOOK_NOTICE_CHOICES.includes(copyright.noticeChoice)||!copyright.noticeText.trim())fail('choose a notice and write or explicitly create its editable text.');
    if(copyright.noticeChoice==='named-license'){
      if(!copyright.licenseName.trim()||!copyright.licenseVersion.trim())fail('enter the named license and version explicitly.');licenseUrl(copyright.licenseUrl);
    }
  }
  return freeze({schemaVersion:1,includeCover:value.includeCover,includeCopyright:value.includeCopyright,cover,copyright});
}
export function workbookFrontMatterRequest(raw){
  if(arguments.length!==1)fail('provide one raw metadata record only.');
  own(raw,['schemaVersion','includeCover','includeCopyright','cover','copyright'],'raw metadata');
  if(raw.schemaVersion!==1)fail('raw metadata version is unsupported.');
  const choice=(value,label)=>{
    own(value,['checked','value','invalid'],label);own(value.invalid,WORKBOOK_FRONT_VALIDITY,`${label} validity`);
    if(typeof value.checked!=='boolean'||typeof value.value!=='string'||value.value.length>1000||WORKBOOK_FRONT_VALIDITY.some(k=>typeof value.invalid[k]!=='boolean'||value.invalid[k]))fail(`${label} has an invalid inclusion choice.`);
    return value.checked;
  };
  const includeCover=choice(raw.includeCover,'include cover'),includeCopyright=choice(raw.includeCopyright,'include copyright');
  const fields=(value,keys,active,label)=>{
    own(value,keys,label);
    const hasLicense=keys.includes('noticeChoice');
    if(hasLicense)own(value.noticeChoice,['value','invalid'],'noticeChoice');
    const licenseActive=hasLicense&&value.noticeChoice.value==='named-license';
    return Object.fromEntries(keys.map(k=>{const entry=own(value[k],['value','invalid'],k);own(entry.invalid,WORKBOOK_FRONT_VALIDITY,`${k} validity`);
      const enabled=active&&(!k.startsWith('license')||licenseActive);
      if(WORKBOOK_FRONT_VALIDITY.some(f=>typeof entry.invalid[f]!=='boolean'||enabled&&entry.invalid[f]))fail(`${k} has unfinished or invalid text; your raw draft is retained.`);
      return [k,entry.value];}));
  };
  return validateWorkbookFrontMatter({schemaVersion:1,includeCover,includeCopyright,
    cover:fields(raw.cover,WORKBOOK_COVER_FIELDS,includeCover,'cover controls'),copyright:fields(raw.copyright,WORKBOOK_COPYRIGHT_FIELDS,includeCopyright,'copyright controls')});
}
const noticeNames=Object.freeze({'all-rights-reserved':'All rights reserved','custom-permissions':'Custom permissions','named-license':'Named license','custom-text':'Custom notice'});
export function workbookCopyrightNoticeStarter(value){
  if(arguments.length!==1)fail('provide one metadata record only.');
  // Validate data/text but do not demand a completed copyright page in order to
  // explicitly request an editable starter. No holder/year/license is inferred.
  own(value,['schemaVersion','includeCover','includeCopyright','cover','copyright'],'metadata');
  if(value.schemaVersion!==1||typeof value.includeCover!=='boolean'||typeof value.includeCopyright!=='boolean')fail('metadata version and inclusion choices are unsupported.');
  const metadata=validateWorkbookFrontMatter({...value,includeCover:false,includeCopyright:false}),c=metadata.copyright;
  if(c.noticeChoice==='all-rights-reserved')return 'All rights reserved.';
  if(c.noticeChoice==='custom-permissions')return 'Specify the uses you authorize, their conditions and restrictions, and how permission can be requested.';
  if(c.noticeChoice==='named-license'){
    if(!c.licenseName.trim()||!c.licenseVersion.trim())fail('enter the named license and version before requesting a starter.');licenseUrl(c.licenseUrl);
    return `This work is offered under ${c.licenseName} ${c.licenseVersion}. License terms: ${c.licenseUrl}.`;
  }
  if(c.noticeChoice==='custom-text')fail('write your custom notice directly; no wording is substituted.');
  fail('choose a notice before requesting a starter.');
}
export function workbookFrontMatterRows(value,organizationName){
  if(arguments.length!==2)fail('provide metadata and the explicit public brand organization only.');
  const metadata=validateWorkbookFrontMatter(value);
  const org=supportedText(organizationName,'public brand organization',120);
  const cover=[],copyright=[],add=(rows,text,style='Normal',extra={})=>rows.push({text,style,...extra});
  if(metadata.includeCover){const c=metadata.cover;
    if(org.trim())add(cover,org,'Meta');add(cover,c.title,'Title');if(c.subtitle.trim())add(cover,c.subtitle,'Subtitle');
    if(c.author.trim())add(cover,`Author: ${c.author}`);if(c.edition.trim())add(cover,`Edition: ${c.edition}`,'Meta');if(c.year.trim())add(cover,`Year: ${c.year}`,'Meta');if(c.audience.trim())add(cover,`Audience: ${c.audience}`);
    cover[cover.length-1].stopKeep=true;
  }
  if(metadata.includeCopyright){const c=metadata.copyright;
    add(copyright,'Copyright','Heading1',{pageBreak:metadata.includeCover});add(copyright,`Rights holder: ${c.rightsHolder}`);add(copyright,`Copyright year: ${c.year}`);
    add(copyright,`Notice choice: ${noticeNames[c.noticeChoice]}`,'Meta');add(copyright,c.noticeText);
    if(c.noticeChoice==='named-license'){add(copyright,`License: ${c.licenseName}`);add(copyright,`License version: ${c.licenseVersion}`);add(copyright,`License terms URL: ${c.licenseUrl}`);}
    if(c.permissionsContact.trim())add(copyright,`Permissions contact: ${c.permissionsContact}`);if(c.credits.trim())add(copyright,`Credits: ${c.credits}`);
  }
  return freeze({schemaVersion:1,includeCover:metadata.includeCover,includeCopyright:metadata.includeCopyright,cover,copyright});
}
