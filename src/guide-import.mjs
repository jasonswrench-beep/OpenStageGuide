// Making a new guide from something outside it (kit slice K7 of docs/KIT-SPEC-2026-11-15.md): a Word document that this program or the Word add-in made, and
// lines of pasted text. Both give the shape the program's project file holds (src/guide-project.mjs): paragraphs with their Word XML, words and style name, and the
// pictures the XML names. Nothing is written over an existing guide: the program opens the result as a new guide beside the others.
//
// A Word file is data a person may be sent. It is read, never opened or run. What it holds that this program does not keep is left out and said so, in
// plain words: numbering and bullets, tracked changes (they are accepted), comments, notes at the foot of a page, headers and footers, and every style the
// guide does not have (those paragraphs become plain). The Word XML of what is kept is cleaned again when the guide opens (src/guide-xml-clean.mjs).
// Pure functions: the parts of a Word file (or text) and an `io` that parses and serialises XML in, a guide out.
import {STYLE_NAMES,STYLE_IDS} from './guide-styles.mjs';
import {para,run} from './guide-ooxml.mjs';
import {cleanGuide,LIMITS,PICTURE_MIME} from './guide-project.mjs';
import {readStyles,textOf} from './guide-check.mjs';
import {BOOKMARK_PREFIX} from './guide-forms.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
/** What a Word file may hold for this import: the whole picture budget is a share of the project file's. */
export const IMPORT_LIMITS=Object.freeze({pictures:LIMITS.pictures,pictureBytes:12*1024*1024,totalPictureBytes:24*1024*1024,pasteLines:5000,pasteLine:2000,title:80});

const kids=el=>Array.from(el?.childNodes||[]).filter(n=>n.nodeType===1);
const isW=(el,name)=>el&&el.namespaceURI===W&&el.localName===name;
const child=(el,name)=>kids(el).find(n=>isW(n,name))||null;
const val=el=>el?(el.getAttributeNS(W,'val')||el.getAttribute('w:val')||''):'';
const clean=t=>String(t??'').replace(/\s+/g,' ').trim();

// ---- styles by name
// The guide's own paragraph styles are found by their names (an add-in or program file names them "OSG Module"), whatever id a file gave them; the built-in ones
// the guide uses are Normal, the six headings, Title and the two levels of the contents.
const OWN=new Map(Object.keys(STYLE_NAMES).map(k=>[STYLE_NAMES[k].toLowerCase(),{name:STYLE_NAMES[k],id:STYLE_IDS[k]}]));
const OWN_IDS=new Set(Object.values(STYLE_IDS));
export function styleFor(rawName){
  const name=clean(rawName),low=name.toLowerCase();
  if(!name||low==='normal')return {name:'Normal',id:null};
  if(OWN.has(low))return OWN.get(low);
  const h=/^heading ([1-6])$/.exec(low);
  if(h)return {name:`heading ${h[1]}`,id:`Heading${h[1]}`};
  if(low==='title')return {name:'Title',id:'Title',builtIn:'Title'};
  const t=/^toc ([1-9])$/.exec(low);
  if(t)return {name:`TOC ${t[1]}`,id:`TOC${t[1]}`,toc:Number(t[1])};
  return null;
}

// The words of a contents line up to its tab: what follows is the page number, which Word works out again.
function textBeforeTab(el){
  let s='',stopped=false;
  const walk=n=>{
    for(const c of kids(n)){
      if(stopped)return;
      if(isW(c,'tab')){stopped=true;return;}
      if(isW(c,'t'))s+=c.textContent;else walk(c);
    }
  };
  walk(el);
  return s;
}

// A paragraph or table element of a Word file, made ready for the guide: its style by name, only the character styles the guide has, tracked changes accepted,
// comments, notes and numbering taken out, and its pictures taken from the file and given the ids the project file names them by.
function prepare(el,ctx){
  const copy=el.cloneNode(true);
  // Tracked changes: what was inserted stays, what was deleted goes.
  const walk=(node,fn)=>{for(const c of kids(node))walk(c,fn);fn(node);};
  const all=[];walk(copy,n=>all.push(n));
  for(const n of all){
    if(!n.parentNode)continue;
    if(isW(n,'del')||isW(n,'moveFrom')){n.parentNode.removeChild(n);ctx.accepted++;}
    else if(isW(n,'ins')||isW(n,'moveTo')){
      while(n.firstChild)n.parentNode.insertBefore(n.firstChild,n);
      n.parentNode.removeChild(n);ctx.accepted++;
    }else if(['commentRangeStart','commentRangeEnd','commentReference','pPrChange','rPrChange','tblPrChange','tcPrChange','trPrChange','sectPrChange','proofErr','permStart','permEnd'].some(t=>isW(n,t))){
      const holder=n.parentNode;holder.removeChild(n);if(/^comment/.test(n.localName))ctx.comments++;
      // A run that only pointed at a comment has nothing left.
      if(isW(holder,'r')&&!kids(holder).some(c=>!isW(c,'rPr')))holder.parentNode?.removeChild(holder);
    }else if(isW(n,'footnoteReference')||isW(n,'endnoteReference')){
      ctx.notes++;const r=n.parentNode;r.parentNode?.removeChild(r);
    }else if(isW(n,'bookmarkStart')||isW(n,'bookmarkEnd')){
      // The guide's own form bookmarks are kept (a form is found by them); every other bookmark is a file's own business.
      if(isW(n,'bookmarkStart')){
        const name=n.getAttributeNS(W,'name')||'';
        if(!name.startsWith(BOOKMARK_PREFIX)){n.parentNode.removeChild(n);ctx.dropped.add(n.getAttributeNS(W,'id'));}
      }else if(ctx.dropped.has(n.getAttributeNS(W,'id')))n.parentNode.removeChild(n);
    }
  }
  for(const pPr of all.filter(n=>isW(n,'pPr')&&n.parentNode)){
    const numPr=child(pPr,'numPr');
    if(numPr){pPr.removeChild(numPr);ctx.numbered++;}
    const sect=child(pPr,'sectPr');
    if(sect){pPr.removeChild(sect);ctx.sections++;}
    const ps=child(pPr,'pStyle');
    if(ps){
      const id=val(ps),named=ctx.styles.get(id)?.name||id,own=styleFor(named);
      if(!own||!own.id){pPr.removeChild(ps);if(!own)ctx.restyled.set(named,(ctx.restyled.get(named)||0)+1);}
      else ps.setAttributeNS(W,'w:val',own.id);
    }
  }
  for(const rs of all.filter(n=>isW(n,'rStyle')&&n.parentNode)){
    const id=val(rs);
    if(!OWN_IDS.has(id))rs.parentNode.removeChild(rs);
  }
  // Pictures: only plain inline pictures of the kinds a Word document holds safely, taken from the file and numbered.
  for(const d of all.filter(n=>isW(n,'drawing')&&n.parentNode)){
    const blip=Array.from(d.getElementsByTagNameNS(A,'blip'))[0];
    const rid=blip?.getAttributeNS(R,'embed')||'';
    const target=rid?ctx.rels.get(rid):null;
    const pic=target?ctx.pictureFor(target):null;
    if(!pic){
      const r=d.parentNode;
      if(isW(r,'r'))r.parentNode?.removeChild(r);else d.parentNode.removeChild(d);
      ctx.leftOut++;
      continue;
    }
    blip.setAttributeNS(R,'r:embed',pic);
  }
  return copy;
}

// The relationships of the document: id -> target path inside the file.
function relationships(relsXml,io){
  const map=new Map();
  if(!relsXml)return map;
  const root=io.parse(relsXml).documentElement;
  for(const r of kids(root)){
    if(r.localName!=='Relationship'||/External/i.test(r.getAttribute('TargetMode')||''))continue;
    const target=r.getAttribute('Target')||'';
    if(!target||/^[a-z][a-z0-9+.-]*:/i.test(target)||target.includes('..'))continue;
    map.set(r.getAttribute('Id'),`word/${target.replace(/^\/?(word\/)?/,'')}`);
  }
  return map;
}
const isPictureBytes=(mime,b64)=>{
  const head=b64.slice(0,16);
  if(mime==='image/png')return head.startsWith('iVBORw0KGgo');
  if(mime==='image/jpeg')return head.startsWith('/9j/');
  if(mime==='image/gif')return head.startsWith('R0lGOD');
  return false;
};

/**
 * A guide from the parts of a Word document: its body, styles, relationships and pictures (`media` maps a path inside the file to {mime,base64}), and the title in
 * its properties. `fallbackTitle` names the guide when the file has none.
 * @param {{document:string, styles?:string|null, rels?:string|null, core?:string|null, media?:Record<string,{mime:string,base64:string}>}} parts
 * @param {{parse:(x:string)=>Document, serialize:(e:Element)=>string}} io
 */
export function guideFromWord(parts,io,{fallbackTitle='Imported guide'}={}){
  if(!parts?.document)throw new Error('That Word file has no document in it.');
  const doc=io.parse(parts.document),root=doc.documentElement;
  if(!root||root.namespaceURI!==W||root.localName!=='document')throw new Error('That file is not a Word document this program can read.');
  const body=child(root,'body');
  if(!body)throw new Error('That Word file has no body to read.');
  const styles=parts.styles?readStyles(io.parse(parts.styles)):new Map();
  const media={},pictureIds=new Map();
  let pictureBytes=0,serial=0;
  const ctx={
    styles,rels:relationships(parts.rels,io),accepted:0,comments:0,notes:0,numbered:0,sections:0,leftOut:0,restyled:new Map(),dropped:new Set(),
    pictureFor(target){
      if(pictureIds.has(target))return pictureIds.get(target);
      const have=parts.media?.[target];
      let id=null;
      if(have&&PICTURE_MIME.includes(have.mime)&&isPictureBytes(have.mime,have.base64)){
        const bytes=Math.floor(have.base64.length*3/4);
        if(serial<IMPORT_LIMITS.pictures&&bytes<=IMPORT_LIMITS.pictureBytes&&pictureBytes+bytes<=IMPORT_LIMITS.totalPictureBytes){
          pictureBytes+=bytes;id=`rIdPicImp${++serial}`;media[id]={mime:have.mime,base64:have.base64};
        }
      }
      pictureIds.set(target,id);
      return id;
    }
  };
  const entries=[];
  const add=(xml,text,style,extra={})=>{
    if(entries.length>=LIMITS.paragraphs)throw new Error('That Word file has more paragraphs than this program can hold.');
    entries.push({xml,text,style,...extra});
  };
  const paragraph=p=>{
    const prepared=prepare(p,ctx);
    // The style it had in the file, by name (prepare has already put the guide's id in, or taken the style off).
    const original=val(child(child(p,'pPr'),'pStyle')),named=styles.get(original)?.name||original,own=styleFor(named);
    const text=textOf(prepared).replace(/\s+$/,'');
    if(own?.toc){
      const words=clean(textBeforeTab(prepared));
      if(own.toc<=2&&words)add('',`${words}\t1`,own.name);
      return;
    }
    // A paragraph that only ended a section, or ended the contents field: nothing is left of it.
    const hasPicture=prepared.getElementsByTagNameNS(W,'drawing').length>0;
    if(!text.trim()&&!hasPicture){
      const onlyFields=Array.from(prepared.getElementsByTagNameNS(W,'r')).every(r=>kids(r).every(c=>isW(c,'rPr')||isW(c,'fldChar')||isW(c,'instrText')));
      const wasSection=!!child(child(p,'pPr'),'sectPr');
      if(wasSection||(onlyFields&&entries.at(-1)?.style?.startsWith('TOC')))return;
    }
    const xml=io.serialize(prepared);
    if(xml.length>LIMITS.paragraphXml){add(para(null,run(text.slice(0,2000))).replace(/^<w:p/,`<w:p xmlns:w="${W}"`),text.slice(0,2000),'Normal');ctx.leftOut++;return;}
    add(xml,text,own?.name||'Normal',own?.builtIn?{builtIn:own.builtIn}:{});
  };
  const table=t=>{
    const prepared=prepare(t,ctx),cells=Array.from(prepared.getElementsByTagNameNS(W,'p'));
    if(!cells.length)return;
    const xml=io.serialize(prepared);
    if(xml.length>LIMITS.paragraphXml){add('',clean(textOf(prepared)),'Normal');ctx.leftOut++;return;}
    cells.forEach((p,i)=>{
      const sid=val(child(child(p,'pPr'),'pStyle')),own=styleFor(styles.get(sid)?.name||sid);
      add(i===0?xml:'',textOf(p).replace(/\s+$/,''),own?.name||'Normal');
    });
  };
  const block=el=>{
    if(isW(el,'p'))paragraph(el);
    else if(isW(el,'tbl'))table(el);
    else if(isW(el,'sdt')){const content=child(el,'sdtContent');if(content)for(const c of kids(content))block(c);}
  };
  for(const el of kids(body))block(el);
  if(!entries.some(e=>e.text.trim()||/<w:drawing/.test(e.xml)))throw new Error('That Word file has no text in it, so there is nothing to make a guide from.');

  // The title: the guide's own title paragraph, else the file's title property, else the name of the file.
  const titled=entries.find(e=>e.style===STYLE_NAMES.guideTitle&&e.text.trim());
  const core=parts.core?clean(Array.from(io.parse(parts.core).getElementsByTagName('dc:title'))[0]?.textContent||''):'';
  const title=clean(titled?.text||core||fallbackTitle).slice(0,IMPORT_LIMITS.title);
  if(!titled)entries.unshift({xml:para(STYLE_IDS.guideTitle,run(title)).replace(/^<w:p/,`<w:p xmlns:w="${W}"`),text:title,style:STYLE_NAMES.guideTitle});

  const modules=entries.filter(e=>e.style===STYLE_NAMES.module).length,lessons=entries.filter(e=>e.style===STYLE_NAMES.lesson).length;
  const notes=[];
  notes.push(`${entries.length===1?'1 paragraph':`${entries.length} paragraphs`} read${modules||lessons?`, with ${modules===1?'1 module':`${modules} modules`} and ${lessons===1?'1 lesson':`${lessons} lessons`}`:''}.`);
  if(!modules&&!lessons)notes.push('The file has no modules or lessons in the guide\'s styles, so everything came in as plain paragraphs. Use Kind of paragraph, or the guide\'s tools, to give it shape.');
  if(Object.keys(media).length)notes.push(`${Object.keys(media).length===1?'1 picture':`${Object.keys(media).length} pictures`} came with it.`);
  if(ctx.leftOut)notes.push(`${ctx.leftOut===1?'1 picture or paragraph':`${ctx.leftOut} pictures or paragraphs`} could not be kept (a kind of picture this program does not hold, one that was too large, or one that was missing from the file).`);
  if(ctx.restyled.size){
    const names=[...ctx.restyled.keys()];
    notes.push(`Paragraphs in styles the guide does not have became plain paragraphs: ${names.slice(0,5).map(n=>`"${n}"`).join(', ')}${names.length>5?` and ${names.length-5} more`:''}. Styles from a Word file (in Format) can bring them in.`);
  }
  if(ctx.numbered)notes.push(`${ctx.numbered===1?'1 paragraph':`${ctx.numbered} paragraphs`} lost their bullets or numbers: the numbering of a Word file is not carried.`);
  if(ctx.accepted)notes.push('Tracked changes were accepted.');
  if(ctx.comments||ctx.notes)notes.push('Comments and notes at the foot of a page are not carried.');
  notes.push('Headers, footers, the look, form numbering and group sizes are not in a Word file, so the guide starts with the program\'s own.');
  return {guide:cleanGuide({title,paragraphs:entries,cursor:0,media}),notes,modules,lessons};
}

// ---- pasted text
const NS=`xmlns:w="${W}"`;
const rooted=xml=>xml.replace(/^<w:(p|tbl)(?=[\s>])/,m=>`${m} ${NS}`);
const MODULE=/^module\s+\d+\s*[:.\-–—]\s*(.+)$/i;
const LESSON=/^lesson\s+\d+\s*\.\s*\d+\s*[:.\-–—]\s*(.+)$/i;
const MINUTES_AT_END=[/\s*[([]\s*(\d{1,3})\s*(?:min(?:ute)?s?)\s*[)\]]\s*$/i,/\s+[-–—]\s*(\d{1,3})\s*(?:min(?:ute)?s?)\s*$/i];
/**
 * A guide from lines of text: "Module 1: Title" and "Lesson 1.1: Title" lines (a lesson may end "(10 min)") become modules and lessons, numbered in order whatever
 * numbers were typed; the first line, when it is neither, is the title; every other line is a plain paragraph.
 * @returns {{guide:any, notes:string[], modules:number, lessons:number}}
 */
export function guideFromText(text,{title=''}={}){
  const lines=String(text??'').replace(/\r\n?/g,'\n').split('\n').map((l,i)=>({text:clean(l),line:i+1})).filter(l=>l.text);
  if(!lines.length)throw new Error('Paste some text first.');
  if(lines.length>IMPORT_LIMITS.pasteLines)throw new Error(`That is more than ${IMPORT_LIMITS.pasteLines} lines. Paste a part of it, or save it as a Word file and import that.`);
  const long=lines.find(l=>l.text.length>IMPORT_LIMITS.pasteLine);
  if(long)throw new Error(`Line ${long.line} is longer than ${IMPORT_LIMITS.pasteLine} letters. Break it into paragraphs first.`);
  const entries=[];
  const add=(xml,words,style)=>entries.push({xml:rooted(xml),text:words,style});
  let name=clean(title),first=true,m=0,l=0,total=0,cumulative=0,loose=0;
  for(const {text:line,line:no} of lines){
    const mod=MODULE.exec(line),les=LESSON.exec(line);
    if(first){
      first=false;
      if(!mod&&!les){
        if(!name)name=line.slice(0,IMPORT_LIMITS.title);
        add(para(STYLE_IDS.guideTitle,run(name)),name,STYLE_NAMES.guideTitle);
        if(name!==line.slice(0,IMPORT_LIMITS.title)&&clean(title)){add(para(null,run(line)),line,'Normal');loose++;}
        continue;
      }
      if(!name)name='Pasted guide';
      add(para(STYLE_IDS.guideTitle,run(name)),name,STYLE_NAMES.guideTitle);
    }
    if(mod){
      m++;l=0;
      const words=`Module ${m}: ${mod[1].trim()}`;
      add(para(STYLE_IDS.module,run(words)),words,STYLE_NAMES.module);
    }else if(les){
      if(!m)throw new Error(`Line ${no} is a lesson before any module. Put a "Module 1: ..." line before it.`);
      l++;total++;
      let rest=les[1].trim(),minutes=null;
      for(const re of MINUTES_AT_END){const hit=re.exec(rest);if(hit){minutes=Number(hit[1]);rest=rest.slice(0,hit.index).trim();break;}}
      if(!rest)throw new Error(`Line ${no} gives a lesson no title.`);
      const words=`Lesson ${m}.${l}: ${rest}`;
      add(para(STYLE_IDS.lesson,run(words)),words,STYLE_NAMES.lesson);
      if(minutes){
        const t=`Time: ${minutes} ${minutes===1?'minute':'minutes'}. Cumulative time: ${cumulative+minutes} ${cumulative+minutes===1?'minute':'minutes'}.`;
        cumulative+=minutes;
        add(para(STYLE_IDS.time,run(t)),t,STYLE_NAMES.time);
      }
    }else{add(para(null,run(line)),line,'Normal');loose++;}
  }
  if(!entries.length)throw new Error('Paste some text first.');
  const notes=[`${m===1?'1 module':`${m} modules`} and ${total===1?'1 lesson':`${total} lessons`} were made from the lines that start "Module" and "Lesson".`];
  if(loose)notes.push(`${loose===1?'1 other line was added as a plain paragraph':`${loose} other lines were added as plain paragraphs`}. Kind of paragraph, or the guide's tools, make ${loose===1?'it a block':'them blocks'}.`);
  if(total&&cumulative)notes.push('Choose Update guide to bring the cumulative times and the numbers up to date.');
  return {guide:cleanGuide({title:name||'Pasted guide',paragraphs:entries,cursor:0}),notes,modules:m,lessons:total};
}
