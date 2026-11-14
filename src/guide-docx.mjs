// Word files made from and read into the guide: the guide as a Word document, the guide's styles as a Word template, and the styles of a Word
// file read so the guide can take them. Pure except for the zip library, which is handed in (JSZip).
//
// The guide is held as a flat Word package (the form Word's own getOoxml gives): the document, its relationships, the styles, the lists and
// the pictures, each as a part. A Word file is the same parts in a zip with a content types list and the package's own relationships.
import {STYLES as FONT_STYLES,obfuscateFont,newGuid,packFontsUsed,namedOnlyUsed} from './guide-fonts.mjs';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT={main:'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',template:'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',settings:'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',fontTable:'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',font:'application/vnd.openxmlformats-officedocument.obfuscatedFont'};
const DECLARATION='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const DOCX_LIMITS=Object.freeze({fileBytes:25*1024*1024,partBytes:6*1024*1024,parts:3000});
/** What a file can carry of fonts: one font file, and all of them together. */
export const EMBED_LIMITS=Object.freeze({fontBytes:3*1024*1024,totalBytes:12*1024*1024});
const SECTION='<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const settingsXml=embed=>`<w:settings xmlns:w="${W}"><w:zoom w:percent="100"/>${embed?'<w:embedTrueTypeFonts/>':''}<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const EXTENSIONS={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',odttf:CT.font};

/**
 * The body of the guide for a Word file: its paragraphs in order, with the contents lines (which the program keeps as plain lines, with
 * no XML) made into a real contents field Word can update to give page numbers.
 * @param {{xml?:string,text?:string,style?:string}[]} entries
 */
export function bodyWithContents(entries){
  let out='',run=[];
  const level=style=>/^toc (\d)$/i.exec(style||'')?.[1];
  const flush=()=>{
    if(!run.length)return;
    run.forEach((e,i)=>{
      const words=String(e.text||'').split('\t')[0].trim();
      const begin=i===0?`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`:'';
      out+=`<w:p><w:pPr><w:pStyle w:val="TOC${level(e.style)}"/></w:pPr>${begin}<w:r><w:t xml:space="preserve">${esc(words)}</w:t></w:r></w:p>`;
    });
    out+='<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
    run=[];
  };
  for(const e of entries){
    if(e.xml&&String(e.xml).trim()){flush();out+=e.xml;}
    else if(level(e.style)&&String(e.text||'').trim())run.push(e);
    // anything else with no XML is a cell of a table that came with the table's own XML
  }
  flush();
  return out;
}

/** Add the section properties a Word page needs, unless the body already ends with some. */
export const withSection=body=>/<w:sectPr[\s>]/.test(body)?body:body+SECTION;

// The two levels of the contents table. Word makes them itself when it updates the table; the entries written here need them to look right until then.
const TOC_STYLES={
  TOC1:'<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>',
  TOC2:'<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="60"/><w:ind w:left="360"/></w:pPr></w:style>',
};
/** The styles part with the contents levels the body uses, when the part does not define them. */
export const withContentsStyles=(stylesXml,body)=>{
  let out=stylesXml;
  for(const [id,def] of Object.entries(TOC_STYLES))
    if(body.includes(`w:val="${id}"`)&&!out.includes(`w:styleId="${id}"`))out=out.replace('</w:styles>',`${def}</w:styles>`);
  return out;
};

const partsOf=(pkgXml,io)=>{
  const root=io.parse(pkgXml).documentElement;
  if(!root||root.localName!=='package')throw new Error('That is not a Word package.');
  const parts=[];
  for(const p of Array.from(root.childNodes).filter(n=>n.nodeType===1&&n.localName==='part')){
    const name=p.getAttribute('pkg:name')||p.getAttributeNS?.('http://schemas.microsoft.com/office/2006/xmlPackage','name');
    const type=p.getAttribute('pkg:contentType')||p.getAttributeNS?.('http://schemas.microsoft.com/office/2006/xmlPackage','contentType');
    const xmlData=Array.from(p.childNodes).find(n=>n.nodeType===1&&n.localName==='xmlData');
    const bin=Array.from(p.childNodes).find(n=>n.nodeType===1&&n.localName==='binaryData');
    if(!name||!type)continue;
    if(bin)parts.push({name,type,base64:(bin.textContent||'').replace(/\s+/g,'')});
    else if(xmlData){const el=Array.from(xmlData.childNodes).find(n=>n.nodeType===1);if(el)parts.push({name,type,xml:io.serialize(el)});}
  }
  return parts;
};

/**
 * A Word file from the flat package: a document (.docx), or, with `template`, a Word template (.dotx) that holds the styles and one empty
 * paragraph. Returns the bytes of the zip. `entries` (the guide's paragraphs) give the body, with the contents lines made into a contents field;
 * without them the body is the package's own.
 * @param {string} pkgXml @param {any} JSZipClass @param {{parse:(x:string)=>Document, serialize:(e:Element)=>string}} io @param {{template?:boolean, entries?:{xml?:string,text?:string,style?:string}[], title?:string, fonts?:(families:string[])=>Promise<Record<string,Partial<Record<'regular'|'bold'|'italic'|'boldItalic',Uint8Array>>>>, onEmbedded?:(list:{family:string,styles:string[],bytes:number}[])=>void, onNamed?:(list:{family:string}[])=>void}} [options]
 *   `fonts` gives the files of the program's own fonts (src/guide-fonts.mjs) by family; the ones the guide uses are carried inside the file.
 *   `onNamed` is told which of the fonts that are only named (no file comes with the program) the guide uses; the file only names them.
 */
export async function docxFromPackage(pkgXml,JSZipClass,io,{template=false,entries,title,fonts,onEmbedded,onNamed}={}){
  const parts=partsOf(pkgXml,io);
  const docPart=parts.find(p=>p.name==='/word/document.xml')?.xml??'';
  const bodyOf=template?'':withSection(entries?bodyWithContents(entries):/<w:body>([\s\S]*)<\/w:body>/.exec(docPart)?.[1]??'');
  onNamed?.(namedOnlyUsed(parts.find(p=>p.name==='/word/styles.xml')?.xml,bodyOf).map(f=>({family:f.family})));
  const zip=new JSZipClass();
  // No entries for folders: some services that open Word files refuse a file that has them.
  const put=(name,data,options)=>zip.file(name,data,{createFolders:false,...options});
  const overrides=[];
  const extensions=new Set();
  let hasSettings=false;
  // The program's own fonts that the guide uses, and their files: carried inside the file, so it looks right where the fonts are not installed.
  const used=typeof fonts==='function'?packFontsUsed(parts.find(p=>p.name==='/word/styles.xml')?.xml,bodyOf):[];
  const loaded=used.length?await fonts(used.map(f=>f.family)):{};
  /** @type {{font:any, files:Record<string,{rid:string,guid:string,target:string}>}[]} */
  const embedded=[];
  let total=0,count=0;
  for(const f of used){
    const have=loaded?.[f.family];
    if(!have)continue;
    /** @type {Record<string,{rid:string,guid:string,target:string}>} */
    const files={};
    for(const style of FONT_STYLES){
      const bytes=have[style];
      if(!(bytes instanceof Uint8Array)||!bytes.length)continue;
      if(bytes.length>EMBED_LIMITS.fontBytes||(total+=bytes.length)>EMBED_LIMITS.totalBytes)throw new Error('The fonts are too large to carry inside a Word file.');
      const guid=newGuid(),n=++count;
      put(`word/fonts/font${n}.odttf`,obfuscateFont(bytes,guid));
      files[style]={rid:`rId${n}`,guid,target:`fonts/font${n}.odttf`};
    }
    if(Object.keys(files).length)embedded.push({font:f,files});
  }
  for(const part of parts){
    const name=part.name.replace(/^\//,'');
    if(!/^[A-Za-z0-9_./-]+$/.test(name)||name.includes('..'))throw new Error('A part of the package has a name a Word file cannot hold.');
    if(part.name==='/word/document.xml'){
      const body=template?`<w:p/>${SECTION}`:bodyOf;
      const root=/^<w:document[^>]*>/.exec(part.xml)?.[0];
      if(!root)throw new Error('The document part of the package is not a document.');
      put(name,`${DECLARATION}${root}<w:body>${body}</w:body></w:document>`);
      overrides.push([part.name,template?CT.template:CT.main]);
    }else if(part.xml!==undefined){
      let xml=part.xml;
      if(part.name==='/word/styles.xml'&&!template)xml=withContentsStyles(xml,bodyOf);
      if(part.name==='/word/_rels/document.xml.rels'&&!/settings\.xml/.test(xml))xml=xml.replace('</Relationships>',`<Relationship Id="rIdSettings" Type="${REL}/settings" Target="settings.xml"/></Relationships>`);
      if(part.name==='/word/_rels/document.xml.rels'&&embedded.length)xml=xml.replace('</Relationships>',`<Relationship Id="rIdFontTable" Type="${REL}/fontTable" Target="fontTable.xml"/></Relationships>`);
      put(name,DECLARATION+xml);
      if(/\.rels$/.test(name)){/* the default for rels covers it */}else overrides.push([part.name,part.type]);
      if(name==='word/settings.xml')hasSettings=true;
    }else if(part.base64!==undefined){
      if(template&&/^word\/media\//.test(name))continue;
      put(name,part.base64,{base64:true});
      const ext=name.split('.').pop().toLowerCase();
      if(EXTENSIONS[ext])extensions.add(ext);
    }
  }
  if(!hasSettings){put('word/settings.xml',DECLARATION+settingsXml(embedded.length>0));overrides.push(['/word/settings.xml',CT.settings]);}
  if(embedded.length){
    const tag={regular:'embedRegular',bold:'embedBold',italic:'embedItalic',boldItalic:'embedBoldItalic'};
    const attr=v=>esc(v).replace(/"/g,'&quot;');
    put('word/fontTable.xml',`${DECLARATION}<w:fonts xmlns:w="${W}" xmlns:r="${REL}">${embedded.map(e=>`<w:font w:name="${attr(e.font.family)}"><w:charset w:val="00"/><w:family w:val="${e.font.kind==='serif'?'roman':'swiss'}"/><w:pitch w:val="variable"/>${FONT_STYLES.filter(k=>e.files[k]).map(k=>`<w:${tag[k]} r:id="${e.files[k].rid}" w:fontKey="${e.files[k].guid}"/>`).join('')}</w:font>`).join('')}</w:fonts>`);
    put('word/_rels/fontTable.xml.rels',`${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${embedded.flatMap(e=>FONT_STYLES.filter(k=>e.files[k]).map(k=>`<Relationship Id="${e.files[k].rid}" Type="${REL}/font" Target="${e.files[k].target}"/>`)).join('')}</Relationships>`);
    overrides.push(['/word/fontTable.xml',CT.fontTable]);
    extensions.add('odttf');
    onEmbedded?.(embedded.map(e=>({family:e.font.family,styles:FONT_STYLES.filter(k=>e.files[k]),bytes:FONT_STYLES.reduce((n,k)=>n+(loaded[e.font.family][k]?.length||0),0)})));
  }
  // The title is a property of the file: a screen reader says it, and Word and PDF files made from this one carry it.
  const heading=String(title??'').replace(/[\u0000-\u001f]/g,' ').trim().slice(0,200);
  if(heading){
    put('docProps/core.xml',`${DECLARATION}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(heading)}</dc:title><dc:language>en-US</dc:language></cp:coreProperties>`);
    overrides.push(['/docProps/core.xml','application/vnd.openxmlformats-package.core-properties+xml']);
  }
  put('_rels/.rels',`${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>${heading?'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>':''}</Relationships>`);
  const types=`${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[...extensions].map(e=>`<Default Extension="${e}" ContentType="${EXTENSIONS[e]}"/>`).join('')}${overrides.map(([n,t])=>`<Override PartName="${esc(n)}" ContentType="${t}"/>`).join('')}</Types>`;
  put('[Content_Types].xml',types);
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE',mimeType:template?'application/vnd.openxmlformats-officedocument.wordprocessingml.template':'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

/**
 * The styles part (and the theme's two fonts) of a Word file, or a message saying why it cannot be read. Nothing in the file is run or opened.
 * @returns {Promise<{stylesXml:string, themeFonts:{major?:string, minor?:string}}>}
 */
export async function readDocxStyles(bytes,JSZipClass,io){
  if(!bytes||!bytes.length||bytes.length>DOCX_LIMITS.fileBytes)throw new Error(`Choose a Word file under ${DOCX_LIMITS.fileBytes/1048576} MB.`);
  let zip;
  try{zip=await JSZipClass.loadAsync(bytes);}catch{throw new Error('That is not a Word file (.docx or .dotx).');}
  const names=Object.keys(zip.files);
  if(names.length>DOCX_LIMITS.parts)throw new Error('That Word file has too many parts to read.');
  const file=zip.file('word/styles.xml');
  if(!file)throw new Error('That file has no styles in it. Choose a Word document or template.');
  if((file._data?.uncompressedSize??0)>DOCX_LIMITS.partBytes)throw new Error('The styles in that file are too large to read.');
  const stylesXml=await file.async('string');
  if(stylesXml.length>DOCX_LIMITS.partBytes)throw new Error('The styles in that file are too large to read.');
  const root=io.parse(stylesXml).documentElement;
  if(!root||root.namespaceURI!==W||root.localName!=='styles')throw new Error('The styles in that file are not in a form this program can read.');
  const themeFonts={};
  const theme=zip.file('word/theme/theme1.xml');
  if(theme&&(theme._data?.uncompressedSize??0)<=DOCX_LIMITS.partBytes){
    const doc=io.parse(await theme.async('string'));
    const A='http://schemas.openxmlformats.org/drawingml/2006/main';
    for(const [key,tag] of [['major','majorFont'],['minor','minorFont']]){
      const latin=Array.from(doc.getElementsByTagNameNS(A,tag)[0]?.childNodes||[]).find(n=>n.nodeType===1&&n.localName==='latin');
      const face=latin?.getAttribute('typeface');
      if(face&&/^[\p{L}\p{N} .,'()+-]{1,64}$/u.test(face))themeFonts[key]=face;
    }
  }
  return {stylesXml,themeFonts};
}
