// Timing & Agenda on the guide: when each lesson (and each timed block in it) starts and ends, worked out from the minutes the guide
// already holds, from a start time the author gives. Nothing is stored: the guide is read from the document each time, the way
// lessons are (guide-model.mjs). The older tool did this from the older course record; this one needs no course record.
//
// Where the minutes come from: a block's own "(N min)" at the end of its first line, and each lesson's "Time:" line. A lesson that has
// minutes on its blocks takes the sum of them (lessonMinutes), as the contents table and Training at a Glance do.
import {readGuide,kindOf,cueFromText,lessonMinutes} from './guide-model.mjs';
import {para,run,esc} from './guide-ooxml.mjs';
import {NS,stylesPart} from './guide-styles.mjs';
import {clockLabel} from './planning.mjs';

export const START_TIME=/^([01]?\d|2[0-3]):([0-5]\d)$/;
const BLOCK_MINUTES=/\s*\((\d+) min\)\s*$/;
const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').trim();

// The blocks of one lesson, in order, each with its cue, its first line and its minutes (0 when it has none).
export function blocksOf(paras,lesson){
  const out=[];
  for(let i=lesson.index+1;i<=lesson.end;i++){
    if(kindOf(paras[i])!=='cue')continue;
    const cue=cueFromText(paras[i].text);
    const raw=clean(paras[i].text).replace(/^[^A-Za-z]+/,'');
    let rest=cue?raw.slice(cue.label.length).replace(/^\s*\t?\s*/,''):raw;
    let minutes=0;
    const m=BLOCK_MINUTES.exec(rest);
    if(m){minutes=Number(m[1]);rest=rest.slice(0,m.index).trim();}
    out.push({index:i,cue:cue?cue.id:'',label:cue?cue.label:'Block',title:rest,minutes});
  }
  return out;
}

// The agenda: rows (module headings, lessons, and the timed blocks in each lesson) with clock times, the total, and notes the author
// should read (a lesson with no minutes; a time line that disagrees with its blocks).
export function guideAgenda(paras,start='09:00'){
  const m=START_TIME.exec(String(start).trim());
  if(!m)throw new Error('Choose a valid start time.');
  const begin=Number(m[1])*3600+Number(m[2])*60;
  const guide=readGuide(paras),rows=[],notes=[];
  let cursor=begin;
  for(const mod of guide.modules){
    if(!mod.lessons.length)continue;
    rows.push({kind:'module',text:mod.text||mod.title,index:mod.index});
    for(const l of mod.lessons){
      const minutes=lessonMinutes(l),first=cursor,last=cursor+minutes*60;
      if(!minutes)notes.push({index:l.index,text:`${l.text} has no minutes, so it takes no time here. Give it a time, or minutes on its blocks.`});
      else if(l.blockMinutes>0&&l.timeMinutes!==null&&l.timeMinutes!==l.blockMinutes)notes.push({index:l.index,text:`${l.text}: its time line says ${l.timeMinutes} minutes but its blocks add up to ${l.blockMinutes}. The agenda uses ${l.blockMinutes}. Choose Update guide to bring them together.`});
      rows.push({kind:'lesson',text:l.text,index:l.index,start:clockLabel(first),end:clockLabel(last),minutes});
      let t=first;
      for(const b of blocksOf(paras,l)){
        if(b.minutes<=0)continue;
        rows.push({kind:'block',text:`${b.label}${b.title?`: ${b.title}`:''}`,cue:b.cue,index:b.index,start:clockLabel(t),end:clockLabel(t+b.minutes*60),minutes:b.minutes});
        t+=b.minutes*60;
      }
      cursor=last;
    }
  }
  const lessons=rows.filter(r=>r.kind==='lesson').length;
  return {title:guide.title||'Guide',start:clockLabel(begin),end:clockLabel(cursor),minutes:(cursor-begin)/60,lessons,rows,notes};
}

// What the pane says above the table.
export const agendaSummary=a=>a.lessons
  ?`${a.lessons} ${a.lessons===1?'lesson':'lessons'}, ${a.minutes} ${a.minutes===1?'minute':'minutes'}, ${a.start} to ${a.end}.`
  :'There are no lessons yet. Add a module and a lesson, then give them minutes.';

// ---- the agenda as a Word file: a table for each module (header row, repeated on each page), lessons in bold, timed blocks indented
const W=[2300,5860,1200];
const text=(t,bold=false)=>`<w:r>${bold?'<w:rPr><w:b/></w:rPr>':''}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const cell=(width,inner)=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const tr=(cells,header=false)=>`<w:tr><w:trPr><w:cantSplit/>${header?'<w:tblHeader/>':''}</w:trPr>${cells.join('')}</w:tr>`;
const BORDER=['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`).join('');
function tableOf(rows){
  const head=tr([cell(W[0],para(null,text('Time',true))),cell(W[1],para(null,text('What',true))),cell(W[2],para(null,text('Minutes',true)))],true);
  const body=rows.map(r=>{
    const bold=r.kind==='lesson',indent=r.kind==='block'?'<w:ind w:left="360"/>':'';
    return tr([cell(W[0],para(null,text(`${r.start} to ${r.end}`,bold))),cell(W[1],para(null,text(r.text,bold),indent)),cell(W[2],para(null,text(String(r.minutes),bold)))]);
  });
  return `<w:tbl><w:tblPr><w:tblW w:w="${W.reduce((a,b)=>a+b)}" w:type="dxa"/><w:tblBorders>${BORDER}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${W.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${head}${body.join('')}</w:tbl>`;
}
export function agendaFiles(a,{language='en-US'}={}){
  const body=[para('Title',run(`${a.title} - agenda`)),para(null,run(`${a.minutes} ${a.minutes===1?'minute':'minutes'} · ${a.start} to ${a.end}`))];
  let group=null;
  const flush=()=>{if(group&&group.rows.length)body.push(para('Heading1',run(group.text)),tableOf(group.rows),para(null,''));};
  for(const r of a.rows){
    if(r.kind==='module'){flush();group={text:r.text,rows:[]};}
    else if(group)group.rows.push(r);
  }
  flush();
  for(const n of a.notes)body.push(para(null,run(`Note: ${n.text}`)));
  const x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  return [
    {name:'[Content_Types].xml',text:`${x}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`},
    {name:'_rels/.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`},
    {name:'word/document.xml',text:`${x}<w:document ${NS}><w:body>${body.join('')}</w:body></w:document>`},
    {name:'word/_rels/document.xml.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`},
    {name:'word/styles.xml',text:`${x}${stylesPart({lang:language},new Set(['Normal','DefaultParagraphFont','Title','Heading1']))}`},
    {name:'docProps/core.xml',text:`${x}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(`${a.title} - agenda`)}</dc:title><dc:language>${esc(language)}</dc:language></cp:coreProperties>`}
  ];
}
export const agendaFileName=title=>`${String(title||'Guide').replace(/[^\p{L}\p{N} _-]/gu,'').trim().slice(0,65)||'Guide'} - agenda.docx`;
