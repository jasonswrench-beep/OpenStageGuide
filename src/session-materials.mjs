// The materials made from the guide for the people who run a session (milestone 2 of the product spec of 2026-11-22): a run sheet for quick delivery and a preparation checklist.
//
// Run sheet: the timeline with clock times, and for each activity what to set up, the supplies, the timed steps with the clock time of each, the debrief prompts, what often goes wrong
// and the other ways to take part (the "offline alternatives" to an interaction on a screen). The facilitator's private notes are in it only when they are asked for.
// Preparation checklist: every supply the activities name, once, with the activities that use it, whether it is printed, brought or opened on a screen, for each person, pair, group
// or room, how many that makes when the guide says how many people and teams there are, the paper, and a status to tick off.
// Both are made as ordinary Word files (the same plain parts as the agenda), as a copy: nothing in the guide changes.
import {guideAgenda} from './guide-agenda.mjs';
import {readBlock} from './guide-model.mjs';
import {para,run,esc,plural} from './guide-ooxml.mjs';
import {NS,stylesPart} from './guide-styles.mjs';
import {readSteps} from './session-text.mjs';
import {activityBlocks} from './session-record.mjs';

const oneLine=t=>String(t??'').replace(/\s+/g,' ').trim();
const lines=t=>String(t??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean);
const clockOf=(base,minutes)=>{const t=base+minutes*60;return `${String(Math.floor(t/3600)%24).padStart(2,'0')}:${String(Math.floor(t%3600/60)).padStart(2,'0')}`;};
const secondsOf=start=>{const m=/^(\d{1,2}):(\d{2})$/.exec(start);return m?Number(m[1])*3600+Number(m[2])*60:0;};

// ---- the run sheet
export function runSheetModel(paras,{start='09:00',lessonIndex=null,includeNotes=false}={}){
  const agenda=guideAgenda(paras,start),lessons=[];
  let current=null;
  for(const r of agenda.rows){
    if(r.kind==='lesson'){current={index:r.index,title:r.text.replace(/\s*\(([^)]*)\)$/,''),start:r.start,end:r.end,minutes:r.minutes,blocks:[]};lessons.push(current);}
    else if(r.kind==='block'&&current)current.blocks.push(blockDetail(paras,r,includeNotes));
  }
  const kept=lessonIndex===null?lessons:lessons.filter(l=>l.index===lessonIndex);
  return {title:agenda.title,start:agenda.start,end:agenda.end,minutes:kept.reduce((n,l)=>n+l.minutes,0),lessons:kept,notes:agenda.notes.map(n=>n.text),includeNotes};
}
function blockDetail(paras,row,includeNotes){
  const base={start:row.start,end:row.end,minutes:row.minutes,cue:row.cue,text:row.text,phase:row.phase||'',title:row.text.replace(/^[^:]+:\s*/,'')};
  if(row.cue!=='activity')return {...base,label:row.text.split(':')[0],detail:false};
  const b=readBlock(paras,row.index),p=b?.parts||{};
  const steps=readSteps(p.steps||'');
  return {...base,label:'Activity',detail:!!b?.parts,grouping:oneLine(p.grouping),output:oneLine(p.output),needs:lines(p.needs),setup:lines(p.preparation),supplies:lines(p.materials),
    steps:steps.steps,otherSteps:steps.without,debrief:lines(p.debrief),challenges:lines(p.challenges),access:lines(p.access),notes:includeNotes?lines(b?.notes):[]};
}

// ---- the Word file's plain parts
const bold=t=>`<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
const cellXml=(w,inner)=>`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
const rowXml=(cells,header=false)=>`<w:tr><w:trPr><w:cantSplit/>${header?'<w:tblHeader/>':''}</w:trPr>${cells.join('')}</w:tr>`;
const BORDER=['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`).join('');
function tableXml(widths,head,rows,caption,description){
  const total=widths.reduce((a,b)=>a+b,0);
  const header=rowXml(head.map((h,i)=>cellXml(widths[i],para(null,bold(h)))),true);
  const body=rows.map(r=>rowXml(r.map((c,i)=>cellXml(widths[i],para(null,run(String(c??'')))))));
  return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblBorders>${BORDER}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCaption w:val="${esc(caption)}"/><w:tblDescription w:val="${esc(description)}"/></w:tblPr><w:tblGrid>${widths.map(w=>`<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${header}${body.join('')}</w:tbl>`;
}
const bullet=t=>para(null,run(t),'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
const labelled=(label,text)=>text?para(null,bold(`${label}: `)+run(text)):'';
const list=(label,items)=>items.length?para(null,bold(`${label}:`),'<w:keepNext/>')+items.map(bullet).join(''):'';
const NUMBERING=`<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
function wordParts(body,{title,language='en-US'}){
  const x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const sect='<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1152" w:bottom="1080" w:left="1152" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  return [
    {name:'[Content_Types].xml',text:`${x}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`},
    {name:'_rels/.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`},
    {name:'word/document.xml',text:`${x}<w:document ${NS}><w:body>${body}${sect}</w:body></w:document>`},
    {name:'word/_rels/document.xml.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`},
    {name:'word/styles.xml',text:`${x}${stylesPart({lang:language},new Set(['Normal','DefaultParagraphFont','Title','Heading1','Heading2']))}`},
    {name:'word/numbering.xml',text:`${x}${NUMBERING}`},
    {name:'docProps/core.xml',text:`${x}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(title)}</dc:title><dc:language>${esc(language)}</dc:language></cp:coreProperties>`}
  ];
}

export function runSheetFiles(model,{language='en-US'}={}){
  const out=[para('Title',run(`Run sheet: ${model.title}`)),para(null,run(`${plural(model.minutes,'minute')}, ${model.start} to ${model.end}.${model.includeNotes?' This copy includes the facilitator\'s private notes.':''}`))];
  if(!model.lessons.length)out.push(para(null,run('There are no lessons to show yet.')));
  for(const lesson of model.lessons){
    out.push(para('Heading1',run(lesson.title)));
    const phased=lesson.blocks.some(b=>b.phase);
    const widths=phased?[2000,4600,1500,1200]:[2000,6100,1200];
    out.push(tableXml(widths,['Time','What',...(phased?['Phase']:[]),'Minutes'],lesson.blocks.map(b=>[`${b.start} to ${b.end}`,b.text,...(phased?[b.phase]:[]),b.minutes]),`Timeline of ${lesson.title}`,'Each row is a block of the lesson with the clock time it starts and ends and its minutes.'));
    out.push(para(null,''));
    for(const b of lesson.blocks.filter(x=>x.detail)){
      out.push(para('Heading2',run(`${b.start} ${b.title} (${b.minutes} min)`)));
      out.push(labelled('Phase',b.phase),labelled('Grouping',b.grouping),labelled('Makes',b.output));
      out.push(list('Needs from earlier work',b.needs),list('Set up',b.setup),list('Supplies',b.supplies));
      if(b.steps.length||b.otherSteps.length){
        let at=secondsOf(b.start.replace(/^Day \d+ · /,''));
        const rows=[...b.steps.map(s=>{const r=[clockOf(at,0),s.min,s.text];at+=s.min*60;return r;}),...b.otherSteps.map(t=>['','',t])];
        out.push(para(null,bold('Steps:'),'<w:keepNext/>'),tableXml([1300,900,7600],['Starts','Minutes','Step'],rows,`Steps of ${b.title}`,'The steps of the activity with the clock time each starts and its minutes.'),para(null,''));
      }
      out.push(list('Debrief prompts',b.debrief),list('If something goes wrong',b.challenges),list('Other ways to take part, such as without a screen',b.access));
      if(b.notes.length)out.push(list('Private notes (facilitator only)',b.notes));
    }
  }
  if(model.notes.length)out.push(list('To look at',model.notes));
  return wordParts(out.join(''),{title:`Run sheet: ${model.title}`,language});
}
export const runSheetFileName=title=>`${String(title||'Guide').replace(/[^\p{L}\p{N} _-]/gu,'').trim().slice(0,60)||'Guide'} - run sheet.docx`;

// ---- the preparation checklist
const PRINT=/\b(card|cards|sheet|sheets|worksheet|handout|checklist|form|ballot|ballots|log|planner|grid|matrix|table|template|record|stickers?|paper|print(?:ed)?|slips?|prompts?|case|scenario|transcript|example|readings?|diagram|forms?)\b/i;
const SCREEN=/\b(slide|slides|digital|online|shared|board|editor|file|files|link|links|website|document|screen)\b/i;
/** Whether a supply is printed, brought or opened on a screen: a word of print wins over a word of screen ("Shared board or worksheet" is printed). */
export function supplyKind(item){
  if(PRINT.test(item))return 'Print';
  if(SCREEN.test(item))return 'Open digitally';
  return 'Bring';
}
const PER_OF={Individual:'person',Pairs:'pair','Small groups':'group','Whole group':'room',Break:'room'};
const MODE_OF=[['Individual',/\b(solo|alone|individual)\b/i],['Pairs',/\bpairs?\b/i],['Small groups',/\b(groups?|teams?|triads?|fours?)\b/i],['Whole group',/\b(whole group|room)\b/i]];
const perFor=grouping=>{
  // "Whole group" is the room, though it has the word "group" in it; and a whole-group share-out at the end does not change who holds the supplies.
  const g=String(grouping||'').replace(/whole group|room/gi,' ');
  // The last way of working named is how many people share an activity's supplies: "Solo then pairs" is pairs; "Solo, pairs, fours, whole group" is groups.
  const last=MODE_OF.filter(([mode])=>mode!=='Whole group').map(([mode,re])=>[mode,g.search(re)]).filter(([,at])=>at>=0).sort((a,b)=>b[1]-a[1])[0];
  return last?PER_OF[last[0]]:'room';
};
const unitsFor=(per,{people=null,teams=null})=>per==='person'?people:per==='pair'?(people?Math.ceil(people/2):null):per==='group'?teams:per==='room'?1:null;
const clip=t=>oneLine(t).replace(/\.$/,'');
const countActivities=n=>`${n} ${n===1?'activity':'activities'}`;

/**
 * Every supply the activities name (their Materials), once, with the activities that use it. `group` is the guide's {people, teams, tables} (the Forms tool's), and `paper` the paper
 * the printed items are for. An item named by several activities is one row; the most that any of them needs is the quantity.
 */
export function prepModel(paras,{lessonIndex=null,group={},paper='Letter'}={}){
  const items=new Map(),activities=activityBlocks(paras,lessonIndex===null?{}:{lessonIndex});
  for(const a of activities){
    const per=perFor(a.parts.grouping||'');
    for(const raw of lines(a.parts.materials)){
      const name=clip(raw),key=name.toLowerCase();
      if(!name)continue;
      const item=items.get(key)||{item:name,kind:supplyKind(name),per,perRank:0,activities:[],quantity:null};
      if(!item.activities.includes(a.title))item.activities.push(a.title);
      const units=unitsFor(per,group);
      if(units!==null&&(item.quantity===null||units>item.quantity)){item.quantity=units;item.per=per;}
      if(!items.has(key))items.set(key,item);
    }
  }
  const rows=[...items.values()].map(({perRank,...r})=>({...r,paper:r.kind==='Print'?paper:'',status:'To do'}));
  const order=['Print','Bring','Open digitally'];
  rows.sort((a,b)=>order.indexOf(a.kind)-order.indexOf(b.kind));
  return {rows,activities:activities.length,counted:group.people!=null||group.teams!=null,paper};
}
export function prepFiles(model,{title='Guide',language='en-US'}={}){
  const out=[para('Title',run(`Preparation checklist: ${title}`))];
  if(!model.rows.length)out.push(para(null,run('No activity names any supplies yet. Add a Materials part to an activity, then make this again.')));
  else out.push(para(null,run(`${plural(model.rows.length,'supply').replace('supplys','supplies')} for ${countActivities(model.activities)}.${model.counted?'The quantities come from the number of people and teams the guide gives.':'The guide does not say how many people and teams there are, so the quantities are left to fill in (the Forms tool takes them).'}`)));
  for(const kind of ['Print','Bring','Open digitally']){
    const rows=model.rows.filter(r=>r.kind===kind);
    if(!rows.length)continue;
    out.push(para('Heading1',run({Print:'To print',Bring:'To bring',['Open digitally']:'To open on a screen'}[kind])));
    out.push(tableXml([2500,2300,900,1100,1000,1000,1100],['Supply','Used in','For each','Quantity','Paper','Status','Notes'],rows.map(r=>[r.item,r.activities.join('; '),r.per,r.quantity??'',r.paper,r.status,'']),`Supplies ${({Print:'to print',Bring:'to bring','Open digitally':'to open on a screen'})[kind]}`,'Each row is a supply, the activities that use it, who gets one, how many, the paper, and its status.'),para(null,''));
  }
  return wordParts(out.join(''),{title:`Preparation checklist: ${title}`,language});
}
export const prepFileName=title=>`${String(title||'Guide').replace(/[^\p{L}\p{N} _-]/gu,'').trim().slice(0,60)||'Guide'} - preparation checklist.docx`;
