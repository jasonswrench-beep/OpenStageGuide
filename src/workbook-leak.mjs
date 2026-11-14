// The leak check: looks at a finished learner workbook (its files, not the code that made it) and says whether anything
// meant for the facilitator is in it. The pane runs it before offering a workbook, and scripts/workbook-leak-check.mjs
// runs the same check on a saved .docx. A pass is NOT Word's Accessibility Checker and NOT proof that the words are right.
// It deliberately does not import the workbook builder: it knows only what a workbook may contain.
import {STYLE_NAMES,STYLE_IDS} from './guide-styles.mjs';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14='http://schemas.microsoft.com/office/word/2010/wordml';
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CORE_REL='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

const ALLOWED_PARTS=new Set(['[Content_Types].xml','_rels/.rels','word/document.xml','word/_rels/document.xml.rels','word/styles.xml','word/settings.xml','word/numbering.xml','word/footer1.xml','docProps/core.xml']);
// A copy Word has saved again carries a few parts of its own. They hold no guide content, but they are still searched for it.
const SAVED_PARTS=['docProps/app.xml','word/theme/theme1.xml','word/fontTable.xml','word/webSettings.xml'];
// Word also writes a footnotes part and an endnotes part into every copy it saves. They are accepted only while they hold nothing but
// Word's own separator lines (found with the real Word of the author's PC, 2026-11-06); a real note, or any text in them, still fails.
const NOTE_PARTS=['word/footnotes.xml','word/endnotes.xml'];
const ALLOWED_MEDIA=/^word\/media\/[^/]+\.png$/;
const ALLOWED_REL_TYPES=[`${REL}/officeDocument`,`${REL}/styles`,`${REL}/settings`,`${REL}/numbering`,`${REL}/image`,`${REL}/footer`,CORE_REL];
const SAVED_REL_TYPES=[`${REL}/extended-properties`,`${REL}/theme`,`${REL}/fontTable`,`${REL}/webSettings`,`${REL}/footnotes`,`${REL}/endnotes`];
// Elements that carry things a workbook must not carry: revisions, comments, notes, hidden text, custom data, legacy pictures
// and text boxes (which hold paragraphs of their own).
const FORBIDDEN_ELEMENTS=['ins','del','moveFrom','moveTo','commentRangeStart','commentRangeEnd','commentReference','footnoteReference','endnoteReference',
  'footnotes','endnotes','comments','customXml','customXmlPr','smartTagPr','pict','object','txbxContent','cellIns','cellDel','cellMerge','ruby'];
const FIELD_OK=/^\s*(TOC|PAGE|NUMPAGES)\b/;
const SAVED_FIELD_OK=/^\s*(TOC|PAGE|NUMPAGES|PAGEREF|HYPERLINK\s+\\l)\b/;
const FACILITATOR_STYLE_NAMES=new Set([STYLE_NAMES.cueFacilitator,STYLE_NAMES.notes,STYLE_NAMES.answerKey,STYLE_NAMES.time,STYLE_NAMES.formTitleFacilitator,STYLE_NAMES.formPrint]);
const FACILITATOR_STYLE_IDS=new Set([STYLE_IDS.cueFacilitator,STYLE_IDS.notes,STYLE_IDS.answerKey,STYLE_IDS.time,STYLE_IDS.formTitleFacilitator,STYLE_IDS.formPrint]);
export const norm=s=>String(s??'').toLowerCase().replace(/[^\p{L}\p{N}]+/gu,'');
const MIN_SEARCH=8; // shorter facilitator texts are matched as whole paragraphs, not searched for inside others

const els=n=>Array.from(n.childNodes).filter(c=>c.nodeType===1);
const all=(n,ns,name)=>Array.from(n.getElementsByTagNameNS(ns,name));
const val=n=>n.getAttributeNS(W,'val')||'';
const textOfParagraph=p=>all(p,W,'t').map(t=>t.textContent).join('');
const everything=root=>[root,...Array.from(root.getElementsByTagName('*'))];

// files: [{name,text}|{name,base64}] (a list or a Map).
// facilitatorTexts: the words of every guide paragraph that is not for learners.
// allowedTexts: what the file may say because the guide's kept content says it (see allowedTextsFrom in guide-workbook.mjs).
//   Words from a facilitator paragraph are a leak only when the file holds them more often than these do, so a learner
//   paragraph that happens to repeat a facilitator one is not refused, but a facilitator paragraph that is added is.
// expectedTitle: what the document's title property should be.
// saved: the file is a copy Word has saved again, so the parts and properties Word adds are accepted.
export function leakCheck({files,facilitatorTexts=[],allowedTexts=[],expectedTitle='',saved=false},{DOMParserClass=globalThis.DOMParser}={}){
  const findings=[],add=(level,rule,message)=>findings.push({level,rule,message});
  const list=Array.isArray(files)?files:[...files].map(([name,v])=>({name,...v}));
  const names=list.map(f=>f.name);
  const okParts=new Set([...ALLOWED_PARTS,...(saved?[...SAVED_PARTS,...NOTE_PARTS]:[])]),okRels=new Set([...ALLOWED_REL_TYPES,...(saved?SAVED_REL_TYPES:[])]);

  // 1. Only the parts a workbook should have.
  const stray=names.filter(n=>!okParts.has(n)&&!ALLOWED_MEDIA.test(n));
  stray.length?add('fail','PARTS',`The file holds parts a workbook should not have: ${stray.join(', ')}.`):add('pass','PARTS','Only the expected parts are in the file.');

  const docs=new Map(),problems=[];
  for(const f of list){
    if(f.text===undefined||!/\.(xml|rels)$/.test(f.name))continue;
    const errors=[],doc=new DOMParserClass({errorHandler:{warning:()=>{},error:e=>errors.push(e),fatalError:e=>errors.push(e)}}).parseFromString(f.text,'application/xml');
    errors.length?problems.push(f.name):docs.set(f.name,doc);
  }
  problems.length?add('fail','XML',`These parts are not well-formed XML: ${problems.join(', ')}.`):add('pass','XML','Every part is well-formed.');

  // 2. No outside links and nothing but the usual relationships.
  const badRels=[];
  for(const [name,doc] of docs)if(name.endsWith('.rels'))
    for(const r of all(doc.documentElement,doc.documentElement.namespaceURI,'Relationship'))
      if(r.getAttribute('TargetMode')==='External'||!okRels.has(r.getAttribute('Type')))badRels.push(`${name}: ${r.getAttribute('Type')||'?'}`);
  badRels.length?add('fail','RELATIONSHIPS',`The file links to something a workbook should not: ${badRels.join('; ')}.`):add('pass','RELATIONSHIPS','No outside links and no unexpected relationships.');

  const document=docs.get('word/document.xml');
  if(!document){add('fail','DOCUMENT','The file has no readable word/document.xml.');return finish(findings);}

  // 3. No facilitator style, in use or defined.
  const used=new Set(all(document.documentElement,W,'pStyle').map(val)),styles=docs.get('word/styles.xml');
  const usedBad=[...used].filter(id=>FACILITATOR_STYLE_IDS.has(id));
  const definedBad=styles?all(styles.documentElement,W,'style').filter(s=>FACILITATOR_STYLE_IDS.has(s.getAttributeNS(W,'styleId'))||FACILITATOR_STYLE_NAMES.has(val(els(s).find(x=>x.localName==='name')||s))).map(s=>s.getAttributeNS(W,'styleId')):[];
  usedBad.length||definedBad.length?add('fail','STYLES',`A facilitator style is in the file: ${[...new Set([...usedBad,...definedBad])].join(', ')}.`):add('pass','STYLES','No facilitator style is used or defined.');

  // 4. Nothing hidden, tracked, commented, boxed or custom.
  const hidden=[],fieldOk=saved?SAVED_FIELD_OK:FIELD_OK;
  for(const [name,doc] of docs){
    if(!name.startsWith('word/')||name.endsWith('.rels')||name.startsWith('word/theme/')||name==='word/fontTable.xml'||name==='word/webSettings.xml')continue;
    if(saved&&NOTE_PARTS.includes(name)){
      // Only Word's separator lines, with no text in them, are the usual empty notes part.
      const one=name.endsWith('footnotes.xml')?'footnote':'endnote';
      const real=all(doc.documentElement,W,one).filter(n=>!['separator','continuationSeparator'].includes(n.getAttributeNS(W,'type')||''));
      if(!real.length&&!all(doc.documentElement,W,'t').length)continue;
      hidden.push(`${name} has ${real.length?`${real.length} real ${one}${real.length===1?'':'s'}`:'text'}`);
      continue;
    }
    for(const el of everything(doc.documentElement)){
      const local=el.localName,isW=el.namespaceURI===W;
      if(isW&&(FORBIDDEN_ELEMENTS.includes(local)||/Change$/.test(local)))hidden.push(`${name} has ${local}`);
      else if(!isW&&(local==='txbxContent'||el.namespaceURI===MC&&local==='AlternateContent'))hidden.push(`${name} has ${local}`);
      else if(isW&&local==='hyperlink'&&!(saved&&el.getAttributeNS(REL,'id')===''&&el.getAttributeNS(W,'anchor')))hidden.push(`${name} has hyperlink`);
      else if(isW&&['vanish','webHidden','specVanish'].includes(local)&&!['false','0','off'].includes(val(el).toLowerCase()))hidden.push(`${name} has hidden text`);
      else if(isW&&local==='sdt'&&!all(el,W14,'checkbox').length)hidden.push(`${name} has a content control that is not a check box`);
      else if(isW&&local==='instrText'&&!fieldOk.test(el.textContent))hidden.push(`${name} has the field ${el.textContent.trim().slice(0,30)}`);
      else if(isW&&local==='fldSimple'&&!fieldOk.test(el.getAttributeNS(W,'instr')))hidden.push(`${name} has a field`);
    }
  }
  hidden.length?add('fail','HIDDEN',`The file holds content a workbook should not: ${[...new Set(hidden)].join('; ')}.`):add('pass','HIDDEN','No comments, tracked changes, hidden text, notes, links, text boxes or custom content controls.');

  // 5. The document properties hold the title and the language and nothing else.
  const core=docs.get('docProps/core.xml');
  if(core){
    const kids=els(core.documentElement),only=['title','language',...(saved?['creator','lastModifiedBy','revision','created','modified','lastPrinted']:[])];
    const extra=kids.filter(k=>!only.includes(k.localName)&&k.textContent.trim()).map(k=>k.localName);
    const title=kids.find(k=>k.localName==='title')?.textContent||'';
    const wrongTitle=expectedTitle&&title!==expectedTitle;
    extra.length||wrongTitle?add('fail','PROPERTIES',extra.length?`The document properties hold more than a title: ${extra.join(', ')}.`:`The document title is "${title}", not "${expectedTitle}".`):add('pass','PROPERTIES','The document properties hold only the title and language.');
  }

  // 6. None of the facilitator paragraphs' words, in the text or in the words that describe it to a screen reader.
  // The words of the body and of any header or footer, which are read by everyone who reads the page.
  const paragraphs=[...docs].filter(([name])=>name==='word/document.xml'||/^word\/(header|footer)\d*\.xml$/.test(name))
    .flatMap(([,doc])=>all(doc.documentElement,W,'p').map(textOfParagraph).map(norm));
  const described=[];
  for(const [name,doc] of docs){
    if(name.endsWith('.rels'))continue;
    for(const el of everything(doc.documentElement)){
      for(const a of ['descr','title'])if(el.getAttribute(a))described.push(norm(el.getAttribute(a)));
      if((el.localName==='docPr'||el.localName==='cNvPr')&&el.getAttribute('name'))described.push(norm(el.getAttribute('name')));
      if(el.namespaceURI===W&&['alias','tag','tblCaption','tblDescription'].includes(el.localName))described.push(norm(val(el)));
    }
  }
  for(const n of ['docProps/core.xml','docProps/app.xml'])if(docs.get(n))described.push(...els(docs.get(n).documentElement).map(k=>norm(k.textContent)));
  const allowed=allowedTexts.map(norm).filter(Boolean);
  const leaked=[];
  for(const text of facilitatorTexts){
    const key=norm(text);
    if(!key)continue;
    let hit;
    if(key.length>=MIN_SEARCH){
      const inFile=paragraphs.filter(p=>p.includes(key)).length,allowedCount=allowed.filter(a=>a.includes(key)).length;
      hit=inFile>allowedCount||(allowedCount===0&&described.some(d=>d.includes(key)));
    }else hit=paragraphs.filter(p=>p===key).length>allowed.filter(a=>a===key).length;
    if(hit)leaked.push(text.length>60?`${text.slice(0,57)}...`:text);
  }
  leaked.length?add('fail','FACILITATOR-TEXT',`Words from ${leaked.length===1?'a facilitator paragraph are':`${leaked.length} facilitator paragraphs are`} in the workbook: "${leaked.slice(0,3).join('", "')}"${leaked.length>3?', and more':''}.`)
    :add('pass','FACILITATOR-TEXT',facilitatorTexts.length?`None of the guide's ${facilitatorTexts.length} facilitator texts are in the file.`:'No facilitator paragraphs were given to search for, so only the structure was checked.');

  // 7. The labels that only a facilitator's paragraphs begin with.
  const labels=paragraphs.filter(p=>p.startsWith('answerkey')||p.startsWith('facilitatornotes'));
  labels.length?add('fail','LABELS','A paragraph begins "Answer key" or "Facilitator notes".'):add('pass','LABELS','No paragraph begins "Answer key" or "Facilitator notes".');
  return finish(findings,facilitatorTexts.length);
}

function finish(findings,searched=0){
  const failed=findings.filter(f=>f.level==='fail');
  return {passed:failed.length===0,findings,failed,searched,
    summary:failed.length?failed.map(f=>f.message).join(' '):'Leak check passed: no facilitator text, answer keys, notes, comments, tracked changes, hidden text or custom content are in the file.'};
}
