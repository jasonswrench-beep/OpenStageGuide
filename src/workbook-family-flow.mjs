import {validateWorkbookLogo} from './workbook-logo.mjs';
import {logoWorkbookBlob} from './workbook-logo-package.mjs';
import {validateWorkbookResearch} from './workbook-research.mjs';
import {validateWorkbookContents} from './workbook-contents.mjs';
import {validateWorkbookFrontMatter} from './workbook-public-frontmatter.mjs';
import {reviewCourse} from './planning.mjs';
import {captureEnvelope} from './checkpoint.mjs';
import {publicOutputText} from './output-text.mjs';
import {resolveSelectedPlanScope} from './selected-plan-scope.mjs';
import {WORKBOOK_FAMILIES,resolveWorkbookBrand,semanticWorkbookFromCourse,semanticWorkbookFromSelected,textContrast} from './workbook-presentation.mjs';
import {FAMILY_WORKBOOK_PARTS,FAMILY_WORKBOOK_SETTINGS_XML,FAMILY_WORKBOOK_CONTENT_TYPES_XML,FAMILY_WORKBOOK_DOCUMENT_RELS_XML,
  WORKBOOK_BASE_LANGUAGES,WORKBOOK_DEFAULT_BASE_LANGUAGE,renderWorkbookFamily} from './workbook-family-output.mjs';

export const FAMILY_BRAND_FIELDS=Object.freeze(['organizationName','text','background','accent','accentText','accentBackground',
  'headingFont','bodyFont','fallbackFont','bodyPt','headingPt','lineSpacing','pageSize','marginIn','responseLineIn','printMode']);
const numbers=new Set(['bodyPt','headingPt','lineSpacing','marginIn','responseLineIn']);
const fail=message=>{throw new Error(`New workbook design: ${message}`);};
const flags=['badInput','customError','patternMismatch','rangeOverflow','rangeUnderflow','stepMismatch','tooLong','tooShort','typeMismatch','valueMissing'];
const own=(value,keys,label)=>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail(`${label} must be a plain data record.`);
  const found=Reflect.ownKeys(value);
  if(found.length!==keys.length||found.some(key=>typeof key!=='string'||!keys.includes(key)))fail(`${label} fields changed.`);
  for(const key of keys){const descriptor=Object.getOwnPropertyDescriptor(value,key);
    if(!descriptor?.enumerable||!Object.hasOwn(descriptor,'value'))fail(`${label}.${key} must be visible own data.`);
  }
  return value;
};

// Every control arrives as raw value plus native validity flags. Empty means
// "use this family's default", so switching families does not copy defaults.
export function familyBrandRequest(raw){
  own(raw,FAMILY_BRAND_FIELDS,'branding controls');
  const request={};
  for(const key of FAMILY_BRAND_FIELDS){
    const entry=own(raw[key],['value','invalid'],key),invalid=own(entry.invalid,flags,`${key} validity`);
    if(typeof entry.value!=='string'||entry.value.length>1000||flags.some(flag=>typeof invalid[flag]!=='boolean'||invalid[flag]))
      fail(`${key} has an unfinished or invalid value; preserve the text and correct it before previewing.`);
    if(entry.value==='')continue;
    if(numbers.has(key)){
      if(!/^(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$/u.test(entry.value))fail(`${key} must be a plain decimal number.`);
      const number=Number(entry.value);
      if(!Number.isFinite(number))fail(`${key} must be a finite number.`);
      request[key]=number;
    }else request[key]=entry.value;
  }
  return request;
}

export function prepareFamilyWorkbook(project,config){
  const hasLanguage=!!config&&typeof config==='object'&&Object.hasOwn(config,'baseLanguage');
  const hasFrontMatter=!!config&&typeof config==='object'&&Object.hasOwn(config,'frontMatter');
  const hasToc=!!config&&typeof config==='object'&&Object.hasOwn(config,'toc');
  const hasResearch=!!config&&typeof config==='object'&&Object.hasOwn(config,'research');
  const hasLogo=!!config&&typeof config==='object'&&Object.hasOwn(config,'logo');
  own(config,['planId','familyId','rawBrand','stamp',...(hasLanguage?['baseLanguage']:[]),...(hasFrontMatter?['frontMatter']:[]),...(hasToc?['toc']:[]),...(hasResearch?['research']:[]),...(hasLogo?['logo']:[])],'review configuration');
  const frontMatter=hasFrontMatter?validateWorkbookFrontMatter(config.frontMatter):null;
  const outputOptions={baseLanguage:hasLanguage?config.baseLanguage:WORKBOOK_DEFAULT_BASE_LANGUAGE,...(hasFrontMatter?{frontMatter}: {}),...(hasToc?{toc:validateWorkbookContents(config.toc)}:{}),...(hasResearch?{research:validateWorkbookResearch(config.research)}:{}),...(hasLogo?{logo:validateWorkbookLogo(config.logo)}:{})};
  const {planId,familyId,rawBrand,stamp}=config;
  const baseLanguage=hasLanguage?config.baseLanguage:WORKBOOK_DEFAULT_BASE_LANGUAGE;
  if(typeof baseLanguage!=='string'||!WORKBOOK_BASE_LANGUAGES.includes(baseLanguage))fail('base language is unsupported; choose a listed document language.');
  if(typeof planId!=='string'||typeof familyId!=='string'||!familyId||typeof stamp!=='string')fail('scope, design or generation changed; refresh the preview.');
  const request=familyBrandRequest(rawBrand),brand=resolveWorkbookBrand(familyId,request);
  if(!brand.eligible)return {eligible:false,request,brand,issues:[...brand.issues],alternative:brand.proposedSafeAlternative};
  if(brand.effective.background.toUpperCase()!=='#FFFFFF')return {eligible:false,request,brand,
    issues:['This generated workbook profile requires a white page background.'],
    alternative:{background:'#FFFFFF',...(textContrast(brand.effective.text,'#FFFFFF')<4.5?{text:'#1A1A1A'}:{})}};
  if(!planId){
    const errors=reviewCourse(project).filter(item=>item.severity==='error');
    if(errors.length)fail('resolve items marked Needs attention in Review Guide before creating a full-course workbook.');
  }
  const semantic=planId?semanticWorkbookFromSelected(project,planId,stamp):semanticWorkbookFromCourse(project);
  const selectedPlanName=planId?resolveSelectedPlanScope(project,planId).planName:null;
  let output;
  try{output=renderWorkbookFamily(semantic,brand,stamp,outputOptions);}
  catch(error){
    if(!/printed text contrast is below 4\.5:1/u.test(error.message||''))throw error;
    for(const alternative of [{text:'#000000'},{accentText:'#000000'},{accentText:'#FFFFFF'},
      {text:'#000000',accentText:'#000000'},{text:'#000000',accentText:'#FFFFFF'}]){
      const candidate=resolveWorkbookBrand(familyId,{...request,...alternative});
      if(!candidate.eligible)continue;
      try{renderWorkbookFamily(semantic,candidate,stamp,outputOptions);
        return {eligible:false,request,brand,issues:['The requested printed text colors fall below 4.5:1 after grayscale conversion.'],alternative};
      }catch{/* Try the next explicit text-color proposal. */}
    }
    return {eligible:false,request,brand,issues:['The requested printed text colors fall below 4.5:1 after grayscale conversion.'],alternative:{}};
  }
  return {eligible:true,request,brand,semantic,selectedPlanName,output};
}

export function familyReviewText(prepared){
  if(!prepared.eligible)fail('an ineligible brand has no workbook preview.');
  const p=prepared.output.preview,b=prepared.brand.effective;
  const role=([name,style])=>`${name}: ${style.font} ${style.sizePt} pt, ${style.align}; space before ${style.beforeTwips/20} pt, after ${style.afterTwips/20} pt${style.fill?`, fill #${style.fill}`:''}${style.border?`, rule #${style.border.color}`:''}`;
  const lines=[
    'NEW GENERATED WORKBOOK COPY — SOURCE PREVIEW ONLY',
    'Page appearance and accessibility have not yet been checked in Word. This includes supported saved learner content only; manual additions to the master are not copied.',
    `Design: ${p.familyName} (version ${p.familyVersion})`,
    `Organization: ${p.organizationName||'(none)'}`,
    `Scope: ${p.scope.label} (${p.scope.kind})`,
    `Project ID: ${p.projectId} · Revision: ${p.revision} · Plan ID: ${p.scope.planId??'(full course)'}`,
    `Generated: ${p.stamp}`,
    `Word compatibility: generated profile ${p.profileVersion} explicitly requests modern Word compatibility mode 15 for this new copy. Word opening behavior and accessibility still require inspection.`,
    `Base document language: ${p.baseLanguage}. Mixed-language passages keep their text but have no authored per-run language tagging in this version.`,
    `Page-flow policy (generated profile ${p.profileVersion}): response label and 1–${p.paginationPolicy.shortResponseMaxLines} writing lines stay together. A prompt may stay with that short area when it has at most ${p.paginationPolicy.shortPromptMaxCharacters} characters, ${p.paginationPolicy.shortPromptMaxExplicitLines} explicit lines, ${p.paginationPolicy.shortPromptMaxEstimatedLines} estimated wrapped lines and a prompt/response group height no greater than ${p.paginationPolicy.shortPromptMaxHeightFraction} of usable page height, and an activity heading of at most ${p.paginationPolicy.shortActivityHeadingMaxEstimatedLines} estimated lines with the whole heading-to-response chain no greater than ${p.paginationPolicy.headingToResponseMaxHeightFraction} of usable page height. The estimate uses page width, font size, indentation and response spacing with a text-line safety factor; the first activity after a module with no objective does not add a Prompt link. Actual Word line breaks can differ. Longer response areas can split; modules do not force a new page.`,
    ...p.identityLines,
    `Requested overrides: ${Object.entries(prepared.brand.rawRequest).map(([key,value])=>`${key} ${value}`).join('; ')||'family defaults'}`,
    `Applied colors: body text ${p.renderedColors.text} on ${p.renderedColors.background}; accent text ${p.renderedColors.accentText} on ${p.renderedColors.accentBackground}; decorative accent ${p.renderedColors.accent}.`,
    `Page: ${p.pageSize}, ${p.marginIn} inch margins, ${p.printMode}; no header, footer or gutter.`,
    `Type: body ${p.appliedStyles.Normal.font} ${p.appliedStyles.Normal.sizePt} pt; heading ${p.appliedStyles.Heading1.font} ${p.appliedStyles.Heading1.sizePt} pt; line spacing ${p.lineSpacing}.`,
    `East Asian/complex-script fallback attribute: ${b.fallbackFont}; actual font installation and Latin substitution are not verified.`,
    'Applied text and response roles:',
    ...Object.entries(p.appliedStyles).map(role),
    `Writing-line height: ${(p.responseLineTwips/1440).toFixed(2)} inches; rule ink #${p.appliedStyles.ResponseLine.border.color}.`,
    'Complete learner content in output order:'
  ];
  if(p.logo)lines.push('Included informative PNG cover logo:',`Alternative description: ${p.logo.alt}`,`PNG pixels: ${p.logo.pixelWidth} by ${p.logo.pixelHeight}; requested width ${p.logo.requestedWidthIn} inches; applied ${p.logo.widthIn.toFixed(4)} by ${p.logo.heightIn.toFixed(4)} inches.`, 'Logo pixels retain their original colors, including in grayscale print mode. Image appearance, alt quality, contrast and native accessibility require review. No decorative-image support is claimed.');
  if(p.frontMatter){
    if(!p.frontMatter.includeCover)lines.push('Existing public opening:',p.title,p.subtitle,...(p.organizationName?[`Organization: ${p.organizationName}`]:[]));
    lines.push(`Public front matter extension: version ${p.frontMatter.schemaVersion}. Cover ${p.frontMatter.includeCover?'included':'off'}; copyright ${p.frontMatter.includeCopyright?'included':'off'}.`, 'Public front matter in exact output order:');
    for(const row of [...p.frontMatter.cover,...p.frontMatter.copyright])lines.push(row.text);
    lines.push('The cover title overrides this copy\'s displayed and file title only. Rights and license choices are authored, not verified or inferred.');
  }
  if(p.toc){
    lines.push(`Contents navigation extension: version ${p.toc.schemaVersion}. Included; internal links only, no calculated page numbers.`,p.toc.note,`Contents entries in heading order (${p.toc.entries.length}):`);
    for(const entry of p.toc.entries)lines.push(`${entry.level===2?'  ':''}${entry.text}`);
    if(!p.toc.entries.length)lines.push('No generated content headings are available.');
  }
  if(p.research){
    lines.push(`Opening educational research extension: version ${p.research.schemaVersion}. Included authored content; claims, references and applicability have not been verified.`,p.research.heading);
    for(const section of p.research.sections)lines.push(section.heading,section.text);
  }
  for(const [mi,module] of p.sections.entries()){
    lines.push(`MODULE ${mi+1}: ${module.displayTitle} · ID ${module.id}`);
    if(module.objective)lines.push('Learning objective:',module.objective);
    for(const [ai,activity] of module.activities.entries()){
      lines.push(`ACTIVITY ${mi+1}.${ai+1}: ${activity.displayTitle} · ID ${activity.id}`,'Participant instructions:',activity.instructions,
        `Writing lines: ${activity.responseLines}${activity.responseLabel?` · ${activity.responseLabel}`:''}`);
    }
  }
  lines.push(p.logo?'Limits: One informative PNG cover logo only; no decorative/JPEG/SVG/background assets, header/footer, arbitrary Word-object transfer, in-place replacement or delivery-pack integration. Inspect layout, image appearance and accessibility before sharing.':'Limits: No logo, header, footer, image, arbitrary Word-object transfer, in-place replacement or delivery-pack integration. Approved font names do not prove installation or Latin fallback behavior. Inspect the new DOCX before sharing.');
  return lines.join('\n');
}

export async function familyWorkbookBlob(files,Zip){
  if(typeof Zip!=='function'||typeof Zip.loadAsync!=='function')fail('the DOCX exporter is unavailable.');
  own(files,FAMILY_WORKBOOK_PARTS,'generated DOCX parts');
  if(FAMILY_WORKBOOK_PARTS.some(name=>typeof files[name]!=='string')||
     files['word/settings.xml']!==FAMILY_WORKBOOK_SETTINGS_XML||
     files['[Content_Types].xml']!==FAMILY_WORKBOOK_CONTENT_TYPES_XML||
     files['word/_rels/document.xml.rels']!==FAMILY_WORKBOOK_DOCUMENT_RELS_XML)fail('generated DOCX compatibility profile changed.');
  // Encode complete strings before JSZip's streaming string-chunk boundaries
  // can split a UTF-16 surrogate pair. Keep canonical expected bytes private:
  // the ZIP implementation receives independent copies, never these snapshots.
  const source={...files},encoder=new TextEncoder();
  const expected=Object.fromEntries(FAMILY_WORKBOOK_PARTS.map(name=>[name,encoder.encode(source[name])]));
  const zip=new Zip();for(const name of FAMILY_WORKBOOK_PARTS)zip.file(name,new Uint8Array(expected[name]),{createFolders:false,binary:true});
  const generated=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  if(!(generated instanceof Uint8Array))fail('the DOCX exporter returned unsupported bytes.');
  // Retain a private copy before the first validation await. Neither the ZIP
  // generator nor the loader receives the buffer that will become the offer.
  const privateBytes=Uint8Array.from(generated);
  exactZipRoster(privateBytes);
  const loaded=await Zip.loadAsync(Uint8Array.from(privateBytes),{checkCRC32:true,createFolders:false}),names=Object.keys(loaded.files);
  if(names.length!==FAMILY_WORKBOOK_PARTS.length||FAMILY_WORKBOOK_PARTS.some(name=>!names.includes(name)))fail('generated DOCX member roster changed.');
  for(const name of FAMILY_WORKBOOK_PARTS){
    const actual=await loaded.file(name)?.async('uint8array'),canonical=expected[name];
    if(!(actual instanceof Uint8Array)||actual.length!==canonical.length||actual.some((byte,i)=>byte!==canonical[i]))fail(`generated DOCX member ${name} changed.`);
  }
  exactZipRoster(privateBytes);
  own(files,FAMILY_WORKBOOK_PARTS,'generated DOCX parts');
  if(FAMILY_WORKBOOK_PARTS.some(name=>files[name]!==source[name]))fail('generated DOCX source strings changed during packaging.');
  return new Blob([Uint8Array.from(privateBytes)],{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

export async function familyWorkbookOutputBlob(output,Zip){
  if(arguments.length!==2)fail('provide generated output and ZIP implementation only.');
  if(!output||typeof output!=='object')fail('generated output is unavailable.');
  const descriptor=Object.getOwnPropertyDescriptor(output,'profileVersion');
  if(!descriptor?.enumerable||!Object.hasOwn(descriptor,'value'))fail('generated output profile changed.');
  if(descriptor.value===4)return logoWorkbookBlob(output,Zip);
  own(output,['profileVersion','familyId','familyVersion','preview','files'],'generated workbook output');
  if(descriptor.value!==3)fail('generated output profile is unsupported.');
  return familyWorkbookBlob(output.files,Zip);
}

function exactZipRoster(bytes){
  // The accepted envelope checker verifies local and central inventories,
  // contiguous payload ranges, names, sizes and CRC fields before JSZip can
  // collapse duplicate names or ignore an unlisted local record.
  if(bytes.length<22)fail('generated DOCX ZIP roster is incomplete.');
  const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),end=bytes.length-22;
  if(view.getUint32(end,true)!==0x06054b50||view.getUint16(end+20,true)!==0)
    fail('generated DOCX ZIP comments or trailing bytes are unavailable.');
  let names;try{names=captureEnvelope(bytes);}catch(error){fail(`generated DOCX ZIP envelope changed: ${error.message}`);}
  if(names.length!==FAMILY_WORKBOOK_PARTS.length||names.some((name,i)=>name!==[...FAMILY_WORKBOOK_PARTS].sort()[i]))
    fail('generated DOCX ZIP member roster changed or duplicated.');
  // Master backups may carry ZIP metadata. A newly generated family workbook
  // has no reason to carry unreviewed entry comments, extras or attributes.
  let at=view.getUint32(end+16,true);
  for(let i=0;i<FAMILY_WORKBOOK_PARTS.length;i++){
    const n=view.getUint16(at+28,true),extra=view.getUint16(at+30,true),comment=view.getUint16(at+32,true);
    const local=view.getUint32(at+42,true);
    if(extra||comment||view.getUint16(local+28,true)||view.getUint16(at+36,true)||view.getUint32(at+38,true)||
       view.getUint16(at+4,true)!==0x14)fail('generated DOCX ZIP entry metadata changed.');
    at+=46+n+extra+comment;
  }
}

const safe=value=>publicOutputText(value).replace(/[^\p{L}\p{N} _-]/gu,'').slice(0,55)||'Workbook';
export function familyWorkbookName(title,familyId,planName=null){
  const family=WORKBOOK_FAMILIES.find(item=>item.id===familyId);
  if(!family)fail('design choice is unavailable.');
  return `${safe(title)} - ${safe(family.name)} Workbook${planName?` - ${safe(planName)}`:''} - ${globalThis.crypto.randomUUID()}.docx`;
}
