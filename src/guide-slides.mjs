// Slides from a presentation, as blocks. The presentation file gives each slide's title, text, speaker notes and number of
// animation clicks (src/pptx.mjs reads it; nothing here draws a slide). The slide pictures come from the author, who exports
// them once from the presentation program (every slide as a picture) and chooses the files here. This file decides which picture
// belongs to which slide and what each slide's block says.
export const MAX_SLIDES_AT_ONCE=40;
export const MAX_TOTAL_PICTURE_BYTES=20*1024*1024;

// The number in a file name: the last run of digits before the extension ("Slide12.PNG" is 12, "deck 3 final 07.jpg" is 7).
export const numberInName=name=>{
  const base=String(name||'').replace(/\.[A-Za-z0-9]+$/,''),m=base.match(/\d+/g);
  return m?Number(m[m.length-1]):null;
};
// Natural order for names without numbers that still sort: "a2" before "a10".
const byName=(a,b)=>String(a.name).localeCompare(String(b.name),'en',{numeric:true,sensitivity:'base'});

// Pairs slide pictures with slides. Files are matched by the number in their name (Slide3.PNG is slide 3), which is how the
// presentation program names an export. When no number can be used, the files are taken in name order, one per slide, and
// only if there are exactly as many as there are slides. Returns {byOrder: Map(slide order -> file), problems: [text]}.
export function matchPictures(slides,files){
  const problems=[],byOrder=new Map(),list=[...files];
  if(!list.length)return {byOrder,problems};
  const nums=list.map(f=>numberInName(f.name));
  const numbered=nums.every(n=>n!==null)&&new Set(nums).size===nums.length;
  if(numbered){
    const orders=new Set(slides.map(s=>s.order));
    const strays=list.filter((f,i)=>!orders.has(nums[i]));
    list.forEach((f,i)=>{if(orders.has(nums[i]))byOrder.set(nums[i],f);});
    if(strays.length)problems.push(`${strays.length===1?'This picture is':'These pictures are'} not for any slide in the presentation (it has ${slides.length}): ${strays.slice(0,5).map(f=>f.name).join(', ')}${strays.length>5?', ...':''}.`);
    return {byOrder,problems};
  }
  if(list.length===slides.length){
    [...list].sort(byName).forEach((f,i)=>byOrder.set(slides[i].order,f));
    return {byOrder,problems};
  }
  problems.push(`The picture file names have no slide numbers and you chose ${list.length} for a presentation of ${slides.length} slides, so they cannot be matched. Export every slide as a picture (File > Export > PNG, all slides) and choose them all again.`);
  return {byOrder,problems};
}

const oneLine=s=>String(s??'').replace(/\s+/g,' ').trim();
// What a screen reader says for a slide's picture: its number and title, and then what is on the slide.
export function slideAlt(slide){
  const title=oneLine(slide.title),text=oneLine(slide.text);
  const on=text&&text!==title?(title&&text.startsWith(title)?text.slice(title.length).trim():text):'';
  const head=`Slide ${slide.order}${title?`: ${title}`:''}.`;
  const all=on?`${head} On the slide: ${on}`:head;
  return all.length>250?`${all.slice(0,247).trim()}...`:all;
}
const clicksLine=n=>n>0?`Animation: ${n} click${n===1?'':'s'}.`:'';

// The words on the slide itself, apart from its title, a line to a line: they go under the slide's picture as "Slide text" in the
// shaded facilitator style, like its speaker notes (author, 2026-11-07).
export function slideTextLines(slide){
  const title=oneLine(slide.title);
  const lines=String(slide.text??'').split(/\r?\n/).map(l=>l.replace(/\s+/g,' ').trim()).filter(Boolean);
  if(lines.length&&title&&lines[0].toLowerCase()===title.toLowerCase())lines.shift();
  return lines.join('\n');
}

// The arguments for one slide's block (see cueBlock): the slide's number and title, its clicks, its picture if it has one,
// and its speaker notes as facilitator notes. `picture` is a picture ready to insert (see image-prepare), or null.
export function slideBlock(slide,{picture=null,share=false}={}){
  const title=oneLine(slide.title);
  const text=[`Slide ${slide.order}${title?`: ${title}`:''}`,clicksLine(slide.clicks)].filter(Boolean).join('\n');
  return {
    cue:'slide',text,audience:share?'everyone':'facilitator',slideText:slideTextLines(slide),
    notes:String(slide.notes??'').split(/\r?\n/).map(l=>l.trim()).filter(Boolean).join('\n'),
    picture:picture?{...picture,alt:slideAlt(slide),decorative:false}:undefined
  };
}
