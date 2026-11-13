import {validateWorkbookLogo,workbookLogoPresentation,workbookLogoBase64,workbookLogoBytes,WORKBOOK_LOGO_MEDIA_PART} from './workbook-logo.mjs';
import {validateWorkbookResearch,workbookResearchRows} from './workbook-research.mjs';
import {validateWorkbookContents,buildWorkbookContents} from './workbook-contents.mjs';
import {validateWorkbookFrontMatter,workbookFrontMatterRows} from './workbook-public-frontmatter.mjs';
import {WORKBOOK_SEMANTIC_VERSION,WORKBOOK_BRAND_POLICY_VERSION,WORKBOOK_MIN_TEXT_CONTRAST,
  WORKBOOK_MAX_SEMANTIC_NODES,WORKBOOK_MAX_SEMANTIC_TEXT_UNITS,resolveWorkbookBrand,workbookFamily,textContrast} from './workbook-presentation.mjs';
import {publicOutputText} from './output-text.mjs';
import {SELECTED_PLAN_LIMITS,SELECTED_WORKBOOK_OBJECTIVES} from './selected-plan-scope.mjs';

// A separate generated-output profile. It does not change the legacy workbook
// producer or the accepted selected-plan delivery-pack version-one validator.
export const FAMILY_WORKBOOK_PROFILE_VERSION=3;
export const WORKBOOK_DEFAULT_BASE_LANGUAGE='en-US';
export const WORKBOOK_BASE_LANGUAGES=Object.freeze(['en-US','en-GB','en-CA','es-ES','es-MX','fr-FR','fr-CA','de-DE']);
export const FAMILY_WORKBOOK_PARTS=Object.freeze(['[Content_Types].xml','_rels/.rels','word/_rels/document.xml.rels',
  'word/document.xml','word/styles.xml','word/settings.xml','docProps/core.xml']);
export const FAMILY_WORKBOOK_SETTINGS_XML='<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
export const FAMILY_WORKBOOK_CONTENT_TYPES_XML='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>';
export const FAMILY_WORKBOOK_DOCUMENT_RELS_XML='<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>';
export const FAMILY_LOGO_PROFILE_VERSION=4;
export const FAMILY_LOGO_ASSET_VERSION=1;
export const FAMILY_LOGO_PARTS=Object.freeze([...FAMILY_WORKBOOK_PARTS,WORKBOOK_LOGO_MEDIA_PART]);
export const FAMILY_LOGO_CONTENT_TYPES_XML=FAMILY_WORKBOOK_CONTENT_TYPES_XML.replace('<Override PartName="/word/document.xml"','<Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml"');
export const FAMILY_LOGO_DOCUMENT_RELS_XML=FAMILY_WORKBOOK_DOCUMENT_RELS_XML.replace('</Relationships>','<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/organization-logo.png"/></Relationships>');
const logoOutputSnapshots=new WeakMap();
export function captureWorkbookLogoOutput(output){
  if(arguments.length!==1)throw new Error('Unsupported logo workbook: provide one owned generated output.');
  const snapshot=output&&typeof output==='object'?logoOutputSnapshots.get(output):null;
  if(!snapshot)throw new Error('Unsupported logo workbook: refresh the current verified logo output; clones and forged output records are unavailable.');
  return Object.freeze({profileVersion:4,assetProfileVersion:1,files:Object.freeze({...snapshot.files}),media:new Uint8Array(snapshot.media)});
}


const refuse=reason=>{throw new Error(`Unsupported styled workbook: ${reason}`);};
const own=(value,keys,label)=>{
  if(!value||typeof value!=='object'||Array.isArray(value)||![Object.prototype,null].includes(Object.getPrototypeOf(value)))refuse(`${label} must be an own-data record.`);
  const found=Reflect.ownKeys(value);
  if(found.length!==keys.length||found.some(k=>typeof k!=='string'||!keys.includes(k)))refuse(`${label} fields are unsupported.`);
  for(const key of keys){const descriptor=Object.getOwnPropertyDescriptor(value,key);
    if(!descriptor||!descriptor.enumerable||!Object.hasOwn(descriptor,'value'))refuse(`${label}.${key} must be an enumerable data field.`);
  }
  return value;
};
const dense=(value,label,max=WORKBOOK_MAX_SEMANTIC_NODES)=>{
  if(!Array.isArray(value)||Object.getPrototypeOf(value)!==Array.prototype||value.length>max)refuse(`${label} is not a bounded dense list.`);
  const keys=Reflect.ownKeys(value);
  if(keys.length!==value.length+1||!keys.includes('length')||keys.some(k=>k!=='length'&&(typeof k!=='string'||!/^(0|[1-9][0-9]*)$/u.test(k)||Number(k)>=value.length)))refuse(`${label} has unsupported fields.`);
  for(let i=0;i<value.length;i++){const d=Object.getOwnPropertyDescriptor(value,String(i));if(!d||!d.enumerable||!Object.hasOwn(d,'value'))refuse(`${label}[${i}] is not data.`);}
  return value;
};
const readable=(value,label,max=50000,required=false)=>{
  if(typeof value!=='string'||value.length>max||(required&&!value.trim()))refuse(`${label} must be supported text.`);
  try{if(publicOutputText(value)!==value)refuse(`${label} contains unsupported output text.`);}catch{refuse(`${label} contains unsupported output text.`);}
  // XML parsers normalize CR and CRLF. Match preview and Word text explicitly.
  return value.replace(/\r\n?/gu,'\n');
};
const id=(value,label)=>readable(value,label,200,true);
const escape=value=>value.replace(/[&<>"']/gu,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const attr=value=>escape(String(value));
const freeze=value=>{if(value&&typeof value==='object'){for(const child of Object.values(value))freeze(child);Object.freeze(value);}return value;};
const same=(a,b)=>{
  if(!a||!b||typeof a!=='object'||typeof b!=='object')return a===b;
  const ka=Object.keys(a).sort(),kb=Object.keys(b).sort();
  return ka.length===kb.length&&ka.every((key,i)=>key===kb[i]&&same(a[key],b[key]));
};
const validateStamp=stamp=>{
  if(typeof stamp!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/u.test(stamp)||
      !Number.isFinite(Date.parse(stamp))||new Date(stamp).toISOString()!==stamp)refuse('generation stamp is invalid.');
  return stamp;
};
const validateLanguageOption=options=>{
  const hasFrontMatter=!!options&&typeof options==='object'&&Object.hasOwn(options,'frontMatter'),hasToc=!!options&&typeof options==='object'&&Object.hasOwn(options,'toc'),hasResearch=!!options&&typeof options==='object'&&Object.hasOwn(options,'research'),hasLogo=!!options&&typeof options==='object'&&Object.hasOwn(options,'logo');
  own(options,['baseLanguage',...(hasFrontMatter?['frontMatter']:[]),...(hasToc?['toc']:[]),...(hasResearch?['research']:[]),...(hasLogo?['logo']:[])],'language option');
  if(typeof options.baseLanguage!=='string'||!WORKBOOK_BASE_LANGUAGES.includes(options.baseLanguage))refuse('base language is unsupported.');
  return options.baseLanguage;
};
const gray=color=>{
  const rgb=[1,3,5].map(i=>parseInt(color.slice(i,i+2),16)/255).map(v=>v<=0.04045?v/12.92:((v+0.055)/1.055)**2.4);
  const luminance=0.2126*rgb[0]+0.7152*rgb[1]+0.0722*rgb[2];
  const channel=luminance<=0.0031308?12.92*luminance:1.055*luminance**(1/2.4)-0.055;
  const hex=Math.round(channel*255).toString(16).padStart(2,'0').toUpperCase();return '#'+hex.repeat(3);
};
function validateSemantic(value){
  own(value,['version','projectId','revision','scope','title','identityLines','nodes'],'semantic workbook');
  if(value.version!==WORKBOOK_SEMANTIC_VERSION)refuse('semantic version is unavailable.');
  const projectId=id(value.projectId,'project ID');
  if(!Number.isSafeInteger(value.revision)||value.revision<0)refuse('revision is invalid.');
  own(value.scope,['kind','planId','label'],'semantic scope');
  if(!['full','selected'].includes(value.scope.kind))refuse('scope is unavailable.');
  if(value.scope.kind==='selected')id(value.scope.planId,'plan ID');
  else if(value.scope.planId!==null)refuse('full scope cannot contain a plan ID.');
  const scopeLabel=readable(value.scope.label,'scope label',300,true);
  const title=readable(value.title,'title',50000,true);
  let textUnits=projectId.length+scopeLabel.length+title.length+(value.scope.planId?.length||0);
  const lines=dense(value.identityLines,'identity lines',100).map(x=>readable(x,'identity line',50000,true));
  if(value.scope.kind==='full'&&lines.length)refuse('full scope cannot carry selected identity lines.');
  if(value.scope.kind==='selected'){
    if(lines.length!==5||!lines[0].startsWith('Delivery plan: ')||!lines[0].endsWith(` · Plan ID: ${value.scope.planId}`)||
       !lines[0].slice('Delivery plan: '.length,-(` · Plan ID: ${value.scope.planId}`).length).trim()||
       !lines[1].startsWith('Exact selected scope: ')||
       !/^Selected planned time: (0|[1-9][0-9]*)(?:\.[0-9]+)? minutes in saved plan order; absolute start times appear only in the agenda\.$/u.test(lines[2])||
       lines[3]!==SELECTED_WORKBOOK_OBJECTIVES||lines[4]!==SELECTED_PLAN_LIMITS)refuse('selected scope identity or policy is unsupported.');
  }
  textUnits+=lines.reduce((n,s)=>n+s.length,0);
  const nodes=dense(value.nodes,'semantic nodes');
  const seen=new Set([projectId]),copy=[];let moduleId=null;
  for(const node of nodes){
    if(!node||typeof node!=='object'||Array.isArray(node)||![Object.prototype,null].includes(Object.getPrototypeOf(node)))refuse('semantic node is unsupported.');
    const kindDescriptor=Object.getOwnPropertyDescriptor(node,'kind');
    if(!kindDescriptor||!kindDescriptor.enumerable||!Object.hasOwn(kindDescriptor,'value'))refuse('semantic node kind is not own data.');
    if(kindDescriptor.value==='module'){
      own(node,['kind','id','title','objective'],'module node');
      const nodeId=id(node.id,'module ID');if(seen.has(nodeId))refuse('node ID is duplicated.');seen.add(nodeId);moduleId=nodeId;
      const heading=readable(node.title,'module title',50000,true),objective=readable(node.objective,'objective');
      textUnits+=nodeId.length+heading.length+objective.length;
      copy.push({kind:'module',id:nodeId,title:heading,objective});
    }else if(kindDescriptor.value==='activity'){
      own(node,['kind','id','moduleId','title','instructions','responseLines'],'activity node');
      if(!moduleId||node.moduleId!==moduleId)refuse('activity order or parent identity changed.');
      const nodeId=id(node.id,'activity ID');if(seen.has(nodeId))refuse('node ID is duplicated.');seen.add(nodeId);
      const heading=readable(node.title,'activity title',50000,true),instructions=readable(node.instructions,'participant instructions');
      if(!Number.isInteger(node.responseLines)||node.responseLines<0||node.responseLines>20)refuse('response-line count is invalid.');
      textUnits+=nodeId.length+heading.length+instructions.length;
      copy.push({kind:'activity',id:nodeId,moduleId,title:heading,instructions,responseLines:node.responseLines});
    }else refuse('semantic node kind is unavailable.');
    if(textUnits>WORKBOOK_MAX_SEMANTIC_TEXT_UNITS)refuse('published semantic text budget exceeded.');
  }
  if(textUnits>WORKBOOK_MAX_SEMANTIC_TEXT_UNITS)refuse('published semantic text budget exceeded.');
  if(value.scope.kind==='selected'){
    if(!copy.some(node=>node.kind==='activity'))refuse('selected scope needs at least one activity.');
    const rosterText=lines[1].slice('Exact selected scope: '.length);
    const roster=(rosterText?rosterText.split(' | '):[]).map(group=>{
      const cut=group.indexOf(': ');if(cut<1)refuse('selected roster is unsupported.');
      return {moduleId:group.slice(0,cut),activityIds:group.slice(cut+2)?group.slice(cut+2).split(', '):[]};
    });
    const actual=[];for(const node of copy){if(node.kind==='module')actual.push({moduleId:node.id,activityIds:[]});else actual.at(-1).activityIds.push(node.id);}
    if(!same(roster.filter(x=>x.activityIds.length),actual)||
       roster.some(x=>!x.moduleId||x.activityIds.some(y=>!y)))refuse('selected roster and published nodes differ.');
  }
  return {projectId,revision:value.revision,scope:{kind:value.scope.kind,planId:value.scope.planId,label:scopeLabel},title,identityLines:[...lines],nodes:copy};
}
function validateBrand(value){
  own(value,['eligible','policyVersion','familyId','familyVersion','rawRequest','effective','issues','proposedSafeAlternative'],'resolved brand');
  if(value.eligible!==true||value.policyVersion!==WORKBOOK_BRAND_POLICY_VERSION)refuse('brand policy is unavailable or unsafe.');
  const family=workbookFamily(value.familyId);
  if(value.familyVersion!==family.version)refuse('family version is unavailable.');
  if(!value.rawRequest||typeof value.rawRequest!=='object'||Array.isArray(value.rawRequest))refuse('raw brand request is unsupported.');
  // The accepted resolver checks all own descriptors and rejects unsupported
  // scalars before exposing an effective brand.
  let expected;try{expected=resolveWorkbookBrand(family.id,value.rawRequest);}catch{refuse('raw brand request is unsupported.');}
  if(!expected.eligible)refuse('brand request is not eligible.');
  own(value.effective,Object.keys(expected.effective),'effective brand');
  dense(value.issues,'brand issues',0);
  own(value.proposedSafeAlternative,[],'brand alternative');
  if(!same(value.effective,expected.effective)||
      !same(value.issues,expected.issues)||!same(value.proposedSafeAlternative,expected.proposedSafeAlternative))refuse('effective brand does not match its reviewed request.');
  if(value.effective.background.toUpperCase()!=='#FFFFFF')refuse('this generated profile requires a white body background.');
  const brand=value.effective.printMode==='grayscale'?{...expected.effective,...Object.fromEntries(['text','background','accent','accentText','accentBackground'].map(key=>[key,gray(expected.effective[key])]))}:{...expected.effective};
  brand.organizationName=brand.organizationName.replace(/\r\n?/gu,'\n');
  if(textContrast(brand.text,brand.background)<WORKBOOK_MIN_TEXT_CONTRAST||
     textContrast(brand.accentText,brand.accentBackground)<WORKBOOK_MIN_TEXT_CONTRAST)refuse('printed text contrast is below 4.5:1.');
  return {family,brand};
}

const recipes=Object.freeze({
  'clearline':{subtitle:'Participant workbook',titleAlign:'left',titleBefore:0,titleAfter:380,moduleBefore:390,activityBefore:220,bodyAfter:150,moduleFill:false,objectiveFill:false,moduleBreak:false,activityPrefix:'',responseLabel:'Response space',lineFactor:1,bodyItalic:false,headingSizeOffset:0,indent:0},
  'editorial-study':{subtitle:'Study workbook',titleAlign:'center',titleBefore:550,titleAfter:260,moduleBefore:470,activityBefore:270,bodyAfter:180,moduleFill:false,objectiveFill:false,moduleBreak:false,activityPrefix:'',responseLabel:'Study notes',lineFactor:1.08,bodyItalic:false,headingSizeOffset:2,indent:170},
  'practice-lab':{subtitle:'Practice workbook',titleAlign:'left',titleBefore:80,titleAfter:230,moduleBefore:320,activityBefore:270,bodyAfter:130,moduleFill:true,objectiveFill:false,moduleBreak:false,activityPrefix:'Practice',responseLabel:'Practice response',lineFactor:1.2,bodyItalic:false,headingSizeOffset:0,indent:100},
  'field-guide':{subtitle:'Field guide workbook',titleAlign:'left',titleBefore:80,titleAfter:180,moduleBefore:250,activityBefore:170,bodyAfter:110,moduleFill:true,objectiveFill:true,moduleBreak:false,activityPrefix:'Activity',responseLabel:'Field notes',lineFactor:1.05,bodyItalic:false,headingSizeOffset:-1,indent:0},
  'reflection-journal':{subtitle:'Reflection workbook',titleAlign:'center',titleBefore:500,titleAfter:330,moduleBefore:420,activityBefore:350,bodyAfter:230,moduleFill:false,objectiveFill:false,moduleBreak:false,activityPrefix:'',responseLabel:'Reflection space',lineFactor:1.45,bodyItalic:false,headingSizeOffset:1,indent:220}
});

const paragraph=(content,style,options={})=>{
  const pPr=`<w:pPr><w:pStyle w:val="${style}"/>${options.keepNext?'<w:keepNext/>':''}${options.pageBreak?'<w:pageBreakBefore/>':''}</w:pPr>`;
  const runs=content.split('\n').map(part=>`<w:r><w:t xml:space="preserve">${escape(part)}</w:t></w:r>`).join('<w:r><w:br/></w:r>');
  const start=options.bookmark?`<w:bookmarkStart w:id="${options.bookmark.bookmarkId}" w:name="${options.bookmark.anchor}"/>`:'',end=options.bookmark?`<w:bookmarkEnd w:id="${options.bookmark.bookmarkId}"/>`:'';
  return `<w:p>${pPr}${start}${runs}${end}</w:p>`;
};
const responseLine=keepNext=>`<w:p><w:pPr><w:pStyle w:val="ResponseLine"/>${keepNext?'<w:keepNext/>':''}</w:pPr><w:r><w:t xml:space="preserve"> </w:t></w:r></w:p>`;
const twipsIn=inches=>Math.round(inches*1440);
const style=(name,{size,font,fallback,color,align='left',before=0,after=0,line=300,outline=null,keep=false,bold=false,italic=false,indent=0,fill=null,border=null,ruleHeight=null})=>{
  // CT_PPr and CT_RPr have ordered children; keep their canonical schema order.
  const pPr=`<w:pPr>${keep?'<w:keepNext/>':''}${border?`<w:pBdr><w:bottom w:val="single" w:sz="${border.size}" w:color="${border.color}"/><w:between w:val="single" w:sz="${border.size}" w:color="${border.color}"/></w:pBdr>`:''}${fill?`<w:shd w:val="clear" w:fill="${fill}"/>`:''}<w:spacing w:before="${before}" w:after="${after}" w:line="${ruleHeight??line}" w:lineRule="${ruleHeight?'exact':'auto'}"/>${indent?`<w:ind w:left="${indent}"/>`:''}<w:jc w:val="${align}"/>${outline!==null?`<w:outlineLvl w:val="${outline}"/>`:''}</w:pPr>`;
  const rPr=`<w:rPr><w:rFonts w:ascii="${attr(font)}" w:hAnsi="${attr(font)}" w:eastAsia="${attr(fallback)}" w:cs="${attr(fallback)}"/>${bold?'<w:b/>':''}${italic?'<w:i/>':''}<w:color w:val="${color}"/><w:sz w:val="${Math.round(size*2)}"/></w:rPr>`;
  return `<w:style w:type="paragraph" w:styleId="${name}" ${name==='Normal'?'w:default="1"':''}><w:name w:val="${name==='Heading1'?'heading 1':name==='Heading2'?'heading 2':name}"/>${name==='Normal'?'':`<w:basedOn w:val="Normal"/>`}${pPr}${rPr}</w:style>`;
};
function styleSpecs(recipe,brand){
  const ink=brand.text.slice(1).toUpperCase(),accent=brand.accent.slice(1).toUpperCase(),accentText=brand.accentText.slice(1).toUpperCase(),accentBackground=brand.accentBackground.slice(1).toUpperCase();
  const bodySize=brand.bodyPt,headingSize=Math.max(14,brand.bodyPt+2,brand.headingPt+recipe.headingSizeOffset),baseLine=Math.round(240*brand.lineSpacing);
  const sectionBorder=recipe.moduleFill?null:{size:8,color:accent};
  const responseRuleColor=textContrast(brand.accent,brand.background)>=3?accent:'767676';
  // Word w:sz stores half-points; the preview reports that applied precision.
  const styled=(name,props)=>({name,props:{fallback:brand.fallbackFont,...props,size:Math.round(props.size*2)/2}});
  return [
    styled('Normal',{size:bodySize,font:brand.bodyFont,color:ink,after:recipe.bodyAfter,line:baseLine}),
    styled('Title',{size:Math.min(36,headingSize+6),font:brand.headingFont,color:ink,align:recipe.titleAlign,before:recipe.titleBefore,after:recipe.titleAfter,keep:true,bold:true}),
    styled('Subtitle',{size:bodySize+2,font:brand.bodyFont,color:ink,align:recipe.titleAlign,after:260,keep:true}),
    styled('Scope',{size:bodySize,font:brand.bodyFont,color:accentText,after:140,line:baseLine,fill:accentBackground,indent:recipe.indent}),
    styled('Meta',{size:Math.max(11,bodySize-0.5),font:brand.bodyFont,color:ink,after:130,line:baseLine}),
    styled('Heading1',{size:headingSize,font:brand.headingFont,color:recipe.moduleFill?accentText:ink,before:recipe.moduleBefore,after:190,outline:0,keep:true,bold:true,fill:recipe.moduleFill?accentBackground:null,border:sectionBorder,indent:recipe.indent}),
    styled('Heading2',{size:headingSize-1,font:brand.headingFont,color:ink,before:recipe.activityBefore,after:130,outline:1,keep:true,bold:true,indent:recipe.indent}),
    styled('Objective',{size:bodySize,font:brand.bodyFont,color:recipe.objectiveFill?accentText:ink,after:210,line:baseLine,italic:recipe.bodyItalic,indent:recipe.indent,fill:recipe.objectiveFill?accentBackground:null}),
    styled('Prompt',{size:bodySize,font:brand.bodyFont,color:ink,after:recipe.bodyAfter,line:baseLine,italic:recipe.bodyItalic,indent:recipe.indent}),
    styled('ResponseLabel',{size:Math.max(11,bodySize-0.5),font:brand.bodyFont,color:accentText,before:100,after:95,line:baseLine,keep:true,bold:true,fill:accentBackground,indent:recipe.indent}),
    styled('ResponseLine',{size:bodySize,font:brand.bodyFont,color:ink,after:25,line:baseLine,ruleHeight:Math.round(twipsIn(brand.responseLineIn)*recipe.lineFactor),border:{size:4,color:responseRuleColor},indent:recipe.indent})
  ];
}
function renderedStyleSpecs(recipe,brand,includeToc=false){
  const styles=styleSpecs(recipe,brand);
  if(includeToc){
    const source=Object.fromEntries(styles.map(({name,props})=>[name,props]));
    styles.push({name:'ContentsTitle',props:{...source.Heading1,outline:0}},{name:'ContentsLevel1',props:{...source.Normal,keep:false,indent:0}},{name:'ContentsLevel2',props:{...source.Normal,keep:false,indent:360}});
  }
  return styles;
}
function stylesXml(recipe,brand,baseLanguage,includeToc=false){
  const styles=renderedStyleSpecs(recipe,brand,includeToc),ink=brand.text.slice(1).toUpperCase(),bodySize=brand.bodyPt;
  return `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${attr(brand.bodyFont)}" w:hAnsi="${attr(brand.bodyFont)}" w:eastAsia="${attr(brand.fallbackFont)}" w:cs="${attr(brand.fallbackFont)}"/><w:color w:val="${ink}"/><w:sz w:val="${Math.round(bodySize*2)}"/><w:lang w:val="${baseLanguage}"/></w:rPr></w:rPrDefault></w:docDefaults>${styles.map(({name,props})=>style(name,props)).join('')}</w:styles>`;
}
function associateShortPrompt(instructions,responseLines,activityHeading,headingCarriesFromModule,styles,page,margin){
  // Heading1 and Heading2 already keep with their next paragraphs. When an
  // empty-objective module leads directly into its first activity, decline an
  // added Prompt link rather than guessing the length of that preceding chain.
  if(responseLines<1||responseLines>3||!instructions.trim()||headingCarriesFromModule)return false;
  const characters=Array.from(instructions);
  if(characters.length>240)return false;
  const explicitLines=instructions.split('\n');
  if(explicitLines.length>3)return false;
  const estimateLines=(content,applied)=>{
    const capacity=Math.floor((page.w-2*margin-(applied.indent??0))/(applied.size*20*0.75));
    if(capacity<1)return Infinity;
    return content.split('\n').reduce((count,line)=>{
      const widthUnits=Array.from(line).reduce((units,character)=>units+(character==='\t'?8:character.codePointAt(0)>0x2e80?2:1),0);
      return count+Math.max(1,Math.ceil(widthUnits/capacity));
    },0);
  };
  const prompt=styles.Prompt,label=styles.ResponseLabel,heading=styles.Heading2,response=styles.ResponseLine;
  const estimatedLines=estimateLines(instructions,prompt),headingLines=estimateLines(activityHeading,heading);
  if(estimatedLines>3||headingLines>2)return false;
  const autoLine=applied=>Math.ceil(applied.size*20*((applied.line??300)/240)*1.2);
  const estimatedHeight=estimatedLines*autoLine(prompt)+Math.max(prompt.after??0,label.before??0)+autoLine(label)+Math.max(label.after??0,response.before??0)+
    responseLines*((response.ruleHeight??0)+(response.after??0));
  const headingHeight=headingLines*autoLine(heading)+(heading.before??0)+(heading.after??0);
  const usableHeight=page.h-2*margin;
  return estimatedHeight<=usableHeight/3&&headingHeight+estimatedHeight<=usableHeight/2;
}
const logoAttribute=value=>escape(value).replace(/\n/gu,'&#10;').replace(/\t/gu,'&#9;');
function logoParagraph(logo,align){
  return `<w:p><w:pPr><w:pStyle w:val="Meta"/><w:keepNext w:val="0"/><w:jc w:val="${align}"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${logo.cx}" cy="${logo.cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Organization logo" descr="${logoAttribute(logo.alt)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Organization logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${logo.cx}" cy="${logo.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}
function bodyAndPreview(semantic,family,brand,stamp,baseLanguage,frontMatter,toc,research,logoChoice){
  const recipe=recipes[family.id];let body='',moduleNumber=0,activityNumber=0;
  const page=brand.pageSize==='Letter'?{w:12240,h:15840}:{w:11906,h:16838},margin=twipsIn(brand.marginIn);
  const specs=renderedStyleSpecs(recipe,brand,!!toc?.include),estimatedStyles=Object.fromEntries(specs.map(({name,props})=>[name,props]));
  const sections=[];let current=null;
  const fragments=[],headings=[];
  const add=(content,style,options)=>{
    if(toc?.include&&(style==='Heading1'||style==='Heading2')){
      const key=options?.headingKey;if(typeof key!=='string')refuse('contents heading identity is unavailable.');
      headings.push({key,level:style==='Heading1'?1:2,text:content});fragments.push({kind:'heading',index:headings.length-1,content,style,options});
    }else fragments.push({kind:'xml',xml:paragraph(content,style,options)});
  };
  const raw=value=>fragments.push({kind:'xml',xml:value});
  const frontRows=frontMatter?workbookFrontMatterRows(frontMatter,brand.organizationName):null;
  const logo=logoChoice?.include?workbookLogoPresentation(logoChoice,{coverIncluded:!!frontRows?.includeCover,pageWidthTwips:page.w,marginTwips:margin}):null;
  const researchRows=research?workbookResearchRows(research):null;
  const frontParagraph=(row,bookmark)=>{
    const pPr=`<w:pPr><w:pStyle w:val="${row.style}"/>${row.stopKeep?'<w:keepNext w:val="0"/>':''}${row.pageBreak?'<w:pageBreakBefore/>':''}</w:pPr>`;
    const runs=row.text.split(/(\n|\t)/u).map(part=>part==='\n'?'<w:r><w:br/></w:r>':part==='\t'?'<w:r><w:tab/></w:r>':`<w:r><w:t xml:space="preserve">${escape(part)}</w:t></w:r>`).join('');
    return `<w:p>${pPr}${bookmark?`<w:bookmarkStart w:id="${bookmark.bookmarkId}" w:name="${bookmark.anchor}"/>`:''}${runs}${bookmark?`<w:bookmarkEnd w:id="${bookmark.bookmarkId}"/>`:''}</w:p>`;
  };
  if(frontRows?.includeCover){if(logo)raw(logoParagraph(logo,recipe.titleAlign));for(const row of frontRows.cover)raw(frontParagraph(row));}
  else{add(semantic.title,'Title');add(recipe.subtitle,'Subtitle');if(brand.organizationName)add(`Organization: ${brand.organizationName}`,'Meta');}
  if(frontRows?.includeCopyright)for(const row of frontRows.copyright){
    if(toc?.include&&row.style==='Heading1'){headings.push({key:'copyright:'+semantic.projectId,level:1,text:row.text});fragments.push({kind:'frontHeading',index:headings.length-1,row});}
    else raw(frontParagraph(row));
  }
  if(toc?.include)fragments.push({kind:'contents'});
  if(researchRows?.include){
    const addResearch=(row,key)=>{
      if(toc?.include&&(row.style==='Heading1'||row.style==='Heading2')){headings.push({key,level:row.style==='Heading1'?1:2,text:row.text});fragments.push({kind:'frontHeading',index:headings.length-1,row});}
      else raw(frontParagraph(row));
    };
    addResearch({text:researchRows.heading,style:'Heading1',pageBreak:!!toc?.include||!!frontRows&&(frontRows.includeCover||frontRows.includeCopyright)},'research:'+semantic.projectId+':title');
    for(const section of researchRows.sections){addResearch({text:section.heading,style:'Heading2'},'research:'+semantic.projectId+':'+section.key);addResearch({text:section.text,style:'Normal'});}
  }
  add(semantic.scope.label,'Scope',toc?.include||researchRows?.include||frontRows&&(frontRows.includeCover||frontRows.includeCopyright)?{pageBreak:true}:undefined);
  add(`Project ID: ${semantic.projectId} · Generation revision: ${semantic.revision}`,'Meta');
  if(semantic.scope.kind==='selected')add(`Plan ID: ${semantic.scope.planId}`,'Meta');
  for(const line of semantic.identityLines)add(line,'Meta');
  add(`Generated: ${stamp}`,'Meta');
  for(const node of semantic.nodes){
    if(node.kind==='module'){
      moduleNumber++;activityNumber=0;
      const displayTitle=recipe.activityPrefix?`Module ${moduleNumber} · ${node.title}`:node.title;
      add(displayTitle,'Heading1',{pageBreak:recipe.moduleBreak&&moduleNumber>1,headingKey:'module:'+node.id});
      if(node.objective){add('Learning objective','Heading2',{headingKey:'objective:'+node.id});add(node.objective,'Objective');}
      current={id:node.id,title:node.title,displayTitle,objective:node.objective,activities:[]};sections.push(current);
    }else{
      activityNumber++;
      const displayTitle=recipe.activityPrefix?`${recipe.activityPrefix} ${moduleNumber}.${activityNumber} · ${node.title}`:node.title;
      add(displayTitle,'Heading2',{headingKey:'activity:'+node.id});
      const shortAnswer=node.responseLines>0&&node.responseLines<=3;
      const shortPrompt=shortAnswer&&associateShortPrompt(node.instructions,node.responseLines,displayTitle,activityNumber===1&&!current.objective,estimatedStyles,page,margin);
      add(node.instructions,'Prompt',{keepNext:shortPrompt});
      if(node.responseLines){add(recipe.responseLabel,'ResponseLabel');for(let i=0;i<node.responseLines;i++)raw(responseLine(shortAnswer&&i<node.responseLines-1));}
      current.activities.push({id:node.id,title:node.title,displayTitle,instructions:node.instructions,responseLines:node.responseLines,responseLabel:node.responseLines?recipe.responseLabel:null});
    }
  }
  const contents=toc?.include?buildWorkbookContents(headings):null;
  const contentsXml=()=>{
    let value=paragraph(contents.title,'ContentsTitle',{pageBreak:!!frontRows&&(frontRows.includeCover||frontRows.includeCopyright)});
    value+=paragraph(contents.note,'Meta');
    if(!contents.entries.length)value+=paragraph('No generated content headings are available.','Normal');
    for(const entry of contents.entries){
      const runs=entry.text.split(/(\n|\t)/u).map(part=>part==='\n'?'<w:r><w:br/></w:r>':part==='\t'?'<w:r><w:tab/></w:r>':`<w:r><w:rPr><w:color w:val="${brand.text.slice(1).toUpperCase()}"/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">${escape(part)}</w:t></w:r>`).join('');
      value+=`<w:p><w:pPr><w:pStyle w:val="ContentsLevel${entry.level}"/></w:pPr><w:hyperlink w:anchor="${entry.anchor}" w:history="1">${runs}</w:hyperlink></w:p>`;
    }
    return value;
  };
  body=fragments.map(fragment=>fragment.kind==='xml'?fragment.xml:fragment.kind==='contents'?contentsXml():fragment.kind==='frontHeading'?frontParagraph(fragment.row,contents.entries[fragment.index]):paragraph(fragment.content,fragment.style,{...fragment.options,bookmark:contents.entries[fragment.index]})).join('');
  body+=`<w:sectPr><w:pgSz w:w="${page.w}" w:h="${page.h}"/><w:pgMar w:top="${margin}" w:bottom="${margin}" w:left="${margin}" w:right="${margin}" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
  const appliedStyles=Object.fromEntries(specs.map(({name,props})=>[name,{font:props.font,fallbackFont:props.fallback,sizePt:props.size,beforeTwips:props.before??0,afterTwips:props.after??0,lineTwips:props.ruleHeight??props.line??300,lineRule:props.ruleHeight?'exact':'auto',align:props.align??'left',indentTwips:props.indent??0,fill:props.fill??null,border:props.border??null,outlineLevel:props.outline??null}]));
  const preview={profileVersion:logo?FAMILY_LOGO_PROFILE_VERSION:FAMILY_WORKBOOK_PROFILE_VERSION,familyId:family.id,familyName:family.name,familyVersion:family.version,
    title:frontRows?.includeCover?frontMatter.cover.title:semantic.title,subtitle:frontRows?.includeCover?frontMatter.cover.subtitle:recipe.subtitle,organizationName:brand.organizationName,scope:semantic.scope,projectId:semantic.projectId,revision:semantic.revision,baseLanguage,
    identityLines:[...semantic.identityLines],stamp,pageSize:brand.pageSize,marginIn:brand.marginIn,
    sectionGeometryTwips:{width:page.w,height:page.h,top:margin,bottom:margin,left:margin,right:margin,header:0,footer:0,gutter:0},printMode:brand.printMode,
    lineSpacing:brand.lineSpacing,effectiveBaseTypography:{bodyPt:brand.bodyPt,headingPt:brand.headingPt},renderedColors:{text:brand.text,background:brand.background,accent:brand.accent,accentText:brand.accentText,accentBackground:brand.accentBackground},
    paginationPolicy:{shortResponseMaxLines:3,shortPromptMaxCharacters:240,shortPromptMaxExplicitLines:3,shortPromptMaxEstimatedLines:3,shortActivityHeadingMaxEstimatedLines:2,shortPromptMaxHeightFraction:'1/3',headingToResponseMaxHeightFraction:'1/2',shortPromptWidthEstimateFactor:0.75,shortPromptAutomaticLineSafetyFactor:1.2,emptyObjectiveFirstActivityPromptLink:false,longResponseSplittable:true,forcedModuleBreaks:false},
    ...(frontRows&&(frontRows.includeCover||frontRows.includeCopyright)?{frontMatter:frontRows}:{}),
    ...(contents?{toc:contents}:{}),
    ...(logo?{logo}:{}),
    ...(researchRows?.include?{research:researchRows}:{}),
    responseLineTwips:Math.round(twipsIn(brand.responseLineIn)*recipe.lineFactor),appliedStyles,sections};
  return {body,preview};
}

export function renderWorkbookFamily(semanticInput,resolvedBrand,stampInput,options){
  const baseLanguage=arguments.length<4?WORKBOOK_DEFAULT_BASE_LANGUAGE:validateLanguageOption(options);
  const semantic=validateSemantic(semanticInput),{family,brand}=validateBrand(resolvedBrand),stamp=validateStamp(stampInput);
  const frontMatter=arguments.length>=4&&Object.hasOwn(options,'frontMatter')?validateWorkbookFrontMatter(options.frontMatter):null;
  const toc=arguments.length>=4&&Object.hasOwn(options,'toc')?validateWorkbookContents(options.toc):null;
  const research=arguments.length>=4&&Object.hasOwn(options,'research')?validateWorkbookResearch(options.research):null;
  const logoChoice=arguments.length>=4&&Object.hasOwn(options,'logo')?validateWorkbookLogo(options.logo):null;
  const {body,preview}=bodyAndPreview(semantic,family,brand,stamp,baseLanguage,frontMatter,toc,research,logoChoice);
  const files={
    '[Content_Types].xml':preview.logo?FAMILY_LOGO_CONTENT_TYPES_XML:FAMILY_WORKBOOK_CONTENT_TYPES_XML,
    '_rels/.rels':'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    'word/_rels/document.xml.rels':preview.logo?FAMILY_LOGO_DOCUMENT_RELS_XML:FAMILY_WORKBOOK_DOCUMENT_RELS_XML,
    'word/document.xml':`<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"${preview.logo?' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"':''}><w:body>${body}</w:body></w:document>`,
    'word/styles.xml':stylesXml(recipes[family.id],brand,baseLanguage,!!toc?.include),
    'word/settings.xml':FAMILY_WORKBOOK_SETTINGS_XML,
    'docProps/core.xml':`<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${frontMatter?.includeCover?escape(frontMatter.cover.title)+(frontMatter.cover.subtitle.trim()?' '+escape(frontMatter.cover.subtitle):''):escape(semantic.title)+' '+escape(recipes[family.id].subtitle)}</dc:title><dc:creator>${frontMatter?.includeCover&&frontMatter.cover.author.trim()?escape(frontMatter.cover.author):'OpenStage Guide'}</dc:creator><dc:language>${baseLanguage}</dc:language><dcterms:created xsi:type="dcterms:W3CDTF">${escape(stamp)}</dcterms:created></cp:coreProperties>`
  };
  if(preview.logo)files[WORKBOOK_LOGO_MEDIA_PART]=workbookLogoBase64(logoChoice.asset);
  const parts=preview.logo?FAMILY_LOGO_PARTS:FAMILY_WORKBOOK_PARTS;
  if(Object.keys(files).length!==parts.length||parts.some(name=>typeof files[name]!=='string'))refuse('generated package profile changed.');
  const output=freeze({profileVersion:preview.logo?FAMILY_LOGO_PROFILE_VERSION:FAMILY_WORKBOOK_PROFILE_VERSION,...(preview.logo?{assetProfileVersion:FAMILY_LOGO_ASSET_VERSION,logoAsset:logoChoice.asset}:{}),familyId:family.id,familyVersion:family.version,preview,files});
  if(preview.logo)logoOutputSnapshots.set(output,{files:Object.freeze({...files}),media:workbookLogoBytes(logoChoice.asset)});
  return output;
}

// Check current generated-page brand rules without consulting learner content.
// Presets use this on load; actual output still validates the complete semantic
// source through renderWorkbookFamily above.
export function validateWorkbookFamilyBrand(resolvedBrand){
  const {family,brand}=validateBrand(resolvedBrand);
  stylesXml(recipes[family.id],brand,WORKBOOK_DEFAULT_BASE_LANGUAGE);
  return Object.freeze({familyId:family.id,familyVersion:family.version,printMode:brand.printMode});
}
