// Designing E-Learning in a guide. A "screen" is a Screen title block and the blocks after it (visual direction, on-screen text,
// narration, interaction, feedback ...) up to the next Screen title, lesson, module or assessment. Nothing here stores anything:
// screens are read from the blocks in the document, the same way lessons are.
import {kindOf,cueFromText,isPictureParagraph} from './guide-model.mjs';
import {cueBlock,para,run,esc} from './guide-ooxml.mjs';
import {NS,stylesPart} from './guide-styles.mjs';

export const WORDS_PER_MINUTE=150; // a steady voice-over pace
// The sample sentences "Add a screen" puts under each part. The review notices any that are still there.
export const PROMPTS=Object.freeze({
  visual:'Describe the picture, animation or layout for this screen.',
  onScreen:'Write the words learners read on this screen.',
  narration:'Write what the voice says, as it will be recorded.',
  interaction:'Say what the learner does (click, drag or choose) and what happens.',
  feedback:'Say what the learner is told after each choice, and why.'
});
const PROMPT_TEXTS=new Set(Object.values(PROMPTS));

const clean=t=>String(t??'').replace(/[\r\n\u0007]+$/,'').trim();
const BLOCK_MINUTES=/\s*\((\d+) min\)\s*$/;
const words=text=>(String(text??'').match(/[\p{L}\p{N}]+(?:['’][\p{L}]+)*/gu)||[]).length;
const heading=p=>/^heading\b/i.test(String(p.style||''))||/^Heading/.test(String(p.builtIn||''));

// The text of the block that starts at paragraph i: its first line after the label and minutes, and the lines under it.
function blockText(paras,i,cue){
  const first=clean(paras[i].text).replace(/^[^A-Za-z]+/,'').slice(cue.label.length).replace(/^\s*\t?\s*/,'').replace(BLOCK_MINUTES,'').trim();
  const lines=first?[first]:[];
  for(let j=i+1;j<paras.length&&kindOf(paras[j])==='body';j++)if(!isPictureParagraph(paras[j])&&clean(paras[j].text))lines.push(clean(paras[j].text));
  return lines.join('\n');
}

// Every screen in the guide, in order: {index, title, lesson, number (within its lesson), blocks:[{cue,index,text}], narration, onScreen,
// narrationWords}. `index` is the paragraph of the Screen title.
export function screensOf(paras){
  const screens=[];
  let lesson='',count=0,cur=null;
  for(let i=0;i<paras.length;i++){
    const p=paras[i],k=kindOf(p);
    if(k==='module'){cur=null;lesson='';count=0;continue;}
    if(k==='lesson'){cur=null;lesson=clean(p.text);count=0;continue;}
    if(k==='assessment'||k==='guide-title'||heading(p)){cur=null;continue;}
    if(k!=='cue')continue;
    const cue=cueFromText(p.text);
    if(!cue)continue;
    if(cue.id==='screen-title'){
      cur={index:i,title:blockText(paras,i,cue),lesson,number:++count,blocks:[]};
      screens.push(cur);
    }else if(cue.id==='transition'){cur=null;continue;}
    if(cur)cur.blocks.push({cue:cue.id,index:i,text:blockText(paras,i,cue)});
  }
  for(const s of screens){
    const join=id=>s.blocks.filter(b=>b.cue===id&&!PROMPT_TEXTS.has(b.text)).map(b=>b.text).filter(Boolean).join('\n');
    s.narration=join('narration');s.onScreen=join('on-screen-text');s.narrationWords=words(s.narration);
  }
  return screens;
}

// How long the voice-over runs, from the number of words.
export const narrationMinutes=n=>Math.round(n/WORDS_PER_MINUTE*10)/10;
export function narrationStats(screens){
  const wordCount=screens.reduce((a,s)=>a+s.narrationWords,0),byLesson=new Map();
  for(const s of screens)byLesson.set(s.lesson,(byLesson.get(s.lesson)||0)+s.narrationWords);
  return {screens:screens.length,words:wordCount,minutes:narrationMinutes(wordCount),
    lessons:[...byLesson].map(([lesson,w])=>({lesson,words:w,minutes:narrationMinutes(w)}))};
}

// Things to look at, in the form the Review screen lists: {level, rule, message, paragraph}. itemNumbers (optional) is the set of
// question numbers the guide's assessment sections hold, so a Knowledge check can be checked against them.
export function elearningIssues(paras,itemNumbers=null){
  const out=[];
  for(const s of screensOf(paras)){
    const name=s.title||`screen ${s.number}`,real=id=>s.blocks.some(b=>b.cue===id&&b.text&&!PROMPT_TEXTS.has(b.text));
    if(!real('narration'))out.push({level:'warn',rule:'EL-NARRATION',message:`The screen "${name}" has no narration yet.`,paragraph:s.index});
    if(!real('on-screen-text'))out.push({level:'warn',rule:'EL-ONSCREEN',message:`The screen "${name}" has no on-screen text yet.`,paragraph:s.index});
    for(const b of s.blocks)if(PROMPT_TEXTS.has(b.text))out.push({level:'warn',rule:'EL-PROMPT',message:`On the screen "${name}", a block still has the sample sentence "${b.text}" Type over it.`,paragraph:b.index});
    s.blocks.forEach((b,n)=>{
      if(b.cue==='interaction'&&!s.blocks.slice(n+1).some(x=>x.cue==='feedback'))
        out.push({level:'warn',rule:'EL-FEEDBACK',message:`The interaction on the screen "${name}" has no Feedback block after it to say what the learner is told.`,paragraph:b.index});
      if(b.cue==='knowledge-check'){
        const m=/Question\s+(\d+)/i.exec(b.text);
        if(!m)out.push({level:'warn',rule:'EL-KC-LINK',message:`The Knowledge check on the screen "${name}" does not say which question it uses. Write, for example, "Question 3".`,paragraph:b.index});
        else if(itemNumbers&&!itemNumbers.has(Number(m[1])))out.push({level:'warn',rule:'EL-KC-MISSING',message:`The Knowledge check on the screen "${name}" uses Question ${m[1]}, and the guide has no such question.`,paragraph:b.index});
      }
    });
  }
  return out;
}

// "Add a screen": the usual run of blocks with a sample sentence under each part, in the order a developer reads them.
export function screenBlocks({title,minutes,interaction=false},ctx){
  const name=String(title??'').trim();
  if(!name)throw new Error('Type a title for the screen first.');
  const block=(cue,text,extra={})=>cueBlock({cue,text,...extra},ctx);
  return [
    block('screen-title',name,minutes?{minutes}:{}),
    block('visual-direction',PROMPTS.visual),
    block('on-screen-text',PROMPTS.onScreen),
    block('narration',PROMPTS.narration),
    ...(interaction?[block('interaction',PROMPTS.interaction),block('feedback',PROMPTS.feedback)]:[])
  ].join('');
}

// ---- the narration script: one file for the voice talent, screen by screen
const lessonsOf=screens=>{const order=[],by=new Map();for(const s of screens){if(!by.has(s.lesson)){by.set(s.lesson,[]);order.push(s.lesson);}by.get(s.lesson).push(s);}return order.map(l=>({lesson:l,screens:by.get(l)}));};
export function scriptText(screens,title){
  const stats=narrationStats(screens),lines=[`${title} - narration script`,`${stats.screens} screens, ${stats.words} words, about ${stats.minutes} minutes at ${WORDS_PER_MINUTE} words a minute.`,''];
  for(const {lesson,screens:list} of lessonsOf(screens)){
    if(lesson)lines.push(lesson,'');
    for(const s of list)lines.push(`Screen ${s.number}: ${s.title||'(no title)'}`,s.narration||'(no narration yet)','');
  }
  return lines.join('\n').trimEnd()+'\n';
}
// The same script as a Word file (the files list zipWorkbook takes). Real headings, so it has a navigation pane.
export function scriptFiles(screens,title,language='en-US'){
  const stats=narrationStats(screens),NSP=NS;
  const body=[
    para('Title',run(`${title} - narration script`)),
    para(null,run(`${stats.screens} screens, ${stats.words} words, about ${stats.minutes} minutes at ${WORDS_PER_MINUTE} words a minute.`))
  ];
  for(const {lesson,screens:list} of lessonsOf(screens)){
    if(lesson)body.push(para('Heading1',run(lesson)));
    for(const s of list){
      body.push(para('Heading2',run(`Screen ${s.number}: ${s.title||'(no title)'}`)));
      const lines=(s.narration||'(no narration yet)').split('\n');
      for(const l of lines)body.push(para(null,run(l)));
    }
  }
  const x='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  return [
    {name:'[Content_Types].xml',text:`${x}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`},
    {name:'_rels/.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`},
    {name:'word/document.xml',text:`${x}<w:document ${NSP}><w:body>${body.join('')}</w:body></w:document>`},
    {name:'word/_rels/document.xml.rels',text:`${x}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`},
    {name:'word/styles.xml',text:`${x}${stylesPart({lang:language},new Set(['Normal','DefaultParagraphFont','Title','Heading1','Heading2']))}`},
    {name:'docProps/core.xml',text:`${x}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(`${title} - narration script`)}</dc:title><dc:language>${esc(language)}</dc:language></cp:coreProperties>`}
  ];
}
export const scriptFileName=(title,ext)=>`${String(title||'Guide').replace(/\s*:\s*/g,' - ').replace(/[\\/*?"<>|]+/g,' ').replace(/\s+/g,' ').trim()||'Guide'} - Narration script.${ext}`;
