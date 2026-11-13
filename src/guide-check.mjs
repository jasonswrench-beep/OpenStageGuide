// Structure rules for a Word document, shared by the pane (which reads the open document) and the command-line
// check (which reads a saved .docx). Both hand in already-parsed XML, so this file needs no parser of its own.
// A pass here is NOT a pass from Word's Accessibility Checker and NOT a screen-reader test.
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const ADEC='http://schemas.microsoft.com/office/drawing/2017/decorative';

const els=n=>Array.from(n.childNodes).filter(c=>c.nodeType===1);
const kids=(n,name,ns=W)=>n?els(n).filter(c=>c.localName===name&&c.namespaceURI===ns):[];
const first=(n,name,ns=W)=>kids(n,name,ns)[0]||null;
const val=(n,name='val')=>n?(n.getAttributeNS(W,name)||''):'';
const all=(n,name,ns=W)=>Array.from(n.getElementsByTagNameNS(ns,name));
const isOn=n=>n&&!['false','0','off'].includes(val(n).toLowerCase());
const rootOf=n=>n?(n.documentElement||n):null;

// Text in reading order. A tab or line break reads as a space, as it does to a screen reader.
export const textOf=n=>{
  let s='';
  const walk=x=>{
    for(const c of Array.from(x.childNodes)){
      if(c.nodeType!==1)continue;
      if(c.namespaceURI===W&&c.localName==='t')s+=c.textContent;
      else if(c.namespaceURI===W&&(c.localName==='tab'||c.localName==='br'))s+=' ';
      else walk(c);
    }
  };
  walk(n);
  return s;
};

export function readStyles(stylesRoot){
  const map=new Map();
  const root=rootOf(stylesRoot);
  if(!root)return map;
  for(const s of all(root,'style')){
    const ppr=first(s,'pPr'),rpr=first(s,'rPr');
    map.set(val(s,'styleId'),{
      type:val(s,'type'),name:val(first(s,'name')),basedOn:val(first(s,'basedOn')),node:s,
      outline:first(ppr,'outlineLvl')?Number(val(first(ppr,'outlineLvl')))+1:0,
      color:val(first(rpr,'color')),fill:val(first(ppr,'shd'),'fill')
    });
  }
  return map;
}
const resolve=(styles,id,pick)=>{
  const seen=new Set();
  while(id&&styles.has(id)&&!seen.has(id)){seen.add(id);const s=styles.get(id),v=pick(s);if(v)return v;id=s.basedOn;}
  return '';
};
export const headingLevel=(styles,id)=>resolve(styles,id,s=>{const m=/^heading (\d)$/i.exec(s.name);return m?Number(m[1]):s.outline;})||0;
export const styleName=(styles,id)=>styles.get(id)?.name||id;
const paragraphStyleId=p=>val(first(first(p,'pPr'),'pStyle'))||'Normal';

const hex6=v=>/^[0-9a-f]{6}$/i.test(v||'')?v:'';
const lum=hex=>{
  const c=[0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)/255).map(v=>v<=0.03928?v/12.92:((v+0.055)/1.055)**2.4);
  return 0.2126*c[0]+0.7152*c[1]+0.0722*c[2];
};
export const contrast=(a,b)=>{const [hi,lo]=[lum(a),lum(b)].sort((x,y)=>y-x);return (hi+0.05)/(lo+0.05);};

const hasPicture=n=>all(n,'drawing').length>0;
const parentOf=(n,name)=>{for(let p=n.parentNode;p;p=p.parentNode)if(p.localName===name&&p.namespaceURI===W)return p;return null;};

// root: the w:document element (or its Document). stylesRoot: the w:styles element or Document, or null.
// title and language come from outside the body: the document properties and the settings part.
export function checkParts({root,stylesRoot,title,language}){
  const doc=rootOf(root),styles=readStyles(stylesRoot);
  const findings=[],add=(level,rule,message,paragraph)=>findings.push(paragraph===undefined?{level,rule,message}:{level,rule,message,paragraph});
  const paras=all(doc,'p'),indexOf=new Map(paras.map((p,i)=>[p,i]));
  const paragraphIndex=n=>{for(let p=n;p;p=p.parentNode)if(indexOf.has(p))return indexOf.get(p);return undefined;};

  const name=(title||'').trim();
  name?add('pass','TITLE',`Document title is set ("${name}").`):add('fail','TITLE','The document has no title in its properties.');

  const stylesEl=rootOf(stylesRoot);
  const defLang=stylesEl&&val(first(first(first(first(stylesEl,'docDefaults'),'rPrDefault'),'rPr'),'lang'));
  (defLang||language)?add('pass','LANGUAGE',`Document language is set (${defLang||language}).`):add('fail','LANGUAGE','No document language is set.');

  let prev=0,count=0,skipped=0;
  paras.forEach((p,i)=>{
    const lvl=headingLevel(styles,paragraphStyleId(p));
    if(!lvl)return;
    count++;
    if(prev===0&&lvl!==1)add('warn','HEADINGS',`The first heading is level ${lvl}, not 1: "${textOf(p).trim().slice(0,50)}".`,i);
    if(prev&&lvl>prev+1){skipped++;add('fail','HEADINGS',`Heading ${lvl} follows heading ${prev}, skipping a level: "${textOf(p).trim().slice(0,50)}".`,i);}
    prev=lvl;
  });
  if(!count)add('fail','HEADINGS','The document has no headings, so the Navigation pane and heading navigation have nothing to use.');
  else if(!skipped)add('pass','HEADINGS',`${count} heading${count===1?'':'s'}; no skipped levels.`);

  all(doc,'tbl').forEach((t,i)=>{
    const label=`Table ${i+1}`,rows=kids(t,'tr'),cols=kids(first(t,'tblGrid'),'gridCol').length,at=paragraphIndex(t.getElementsByTagNameNS(W,'p')[0]||t);
    add('info','TABLE',`${label}: ${cols} columns, ${rows.length} rows.`);
    isOn(first(first(rows[0],'trPr'),'tblHeader'))?add('pass','TABLE-HEADER',`${label} has a header row.`):add('fail','TABLE-HEADER',`${label} has no header row marked.`,at);
    if(parentOf(t,'tc'))add('fail','TABLE-NESTED',`${label} is nested inside another table.`,at);
    let merged=0,blank=0;
    for(const tc of all(t,'tc')){
      const pr=first(tc,'tcPr');
      if((first(pr,'gridSpan')&&Number(val(first(pr,'gridSpan')))>1)||first(pr,'vMerge')||first(pr,'hMerge'))merged++;
      if(!textOf(tc).trim()&&!hasPicture(tc))blank++;
    }
    if(merged)add('warn','TABLE-MERGED',`${label} has ${merged} merged or split cell${merged===1?'':'s'}.`,at);
    if(blank)add('warn','TABLE-BLANK',`${label} has ${blank} blank cell${blank===1?'':'s'}.`,at);
  });

  const pics=all(doc,'docPr',WP);
  let withAlt=0,decorative=0,missing=0,firstMissing;
  for(const d of pics){
    const dec=all(d,'decorative',ADEC).some(x=>['1','true'].includes(x.getAttribute('val')));
    const alt=(d.getAttribute('descr')||'').trim();
    if(dec)decorative++;
    else if(alt){withAlt++;if(/\.(png|jpe?g|gif|svg)$/i.test(alt))add('warn','IMAGE-ALT',`Alt text looks like a file name: "${alt}".`,paragraphIndex(d));}
    else{missing++;if(firstMissing===undefined)firstMissing=paragraphIndex(d);}
  }
  if(pics.length){
    missing?add('fail','IMAGE-ALT',`${missing} of ${pics.length} pictures have neither alt text nor a decorative mark.`,firstMissing):add('pass','IMAGE-ALT',`${pics.length} pictures: ${withAlt} with alt text, ${decorative} marked decorative.`);
  }
  const anchors=all(doc,'anchor',WP);
  if(anchors.length)add('warn','IMAGE-FLOATING',`${anchors.length} picture${anchors.length===1?' is':'s are'} floating, not in line with text.`,paragraphIndex(anchors[0]));

  const boxes=all(doc,'txbxContent'),frames=all(doc,'framePr');
  if(boxes.length)add('warn','TEXT-BOX',`${boxes.length} text box${boxes.length===1?'':'es'}: reading order can be unpredictable.`,paragraphIndex(boxes[0]));
  if(frames.length)add('warn','FRAME',`${frames.length} paragraph frame${frames.length===1?'':'s'}: reading order can be unpredictable.`,paragraphIndex(frames[0]));

  const typed=[];
  paras.forEach((p,i)=>{
    if(first(first(p,'pPr'),'numPr'))return;
    if(/^\s*(?:[•●▪\-*]|\d+[.)]|[A-Za-z][.)])\s/.test(textOf(p)))typed.push(i);
  });
  if(typed.length)add('warn','TYPED-LIST',`${typed.length} paragraph${typed.length===1?' looks':'s look'} like a typed list (use Word's list tools instead).`,typed[0]);

  let streak=0,streaks=0,streakAt;
  const flush=at=>{if(streak>=2){streaks++;if(streakAt===undefined)streakAt=at;}streak=0;};
  // Empty paragraphs at the very end of the document are Word's own last paragraph (and the one its export adds after
  // it), not spacing someone typed, so a run that reaches the end of the body is not counted.
  const walk=(n,top=false)=>{
    let start;
    for(const c of els(n)){
      if(c.localName==='p'&&c.namespaceURI===W){
        // An answer line is an empty paragraph on purpose: a line to write on, not spacing someone typed.
        const writingLine=val(first(first(c,'pPr'),'pStyle'))==='OSGAnswerLine';
        if(!textOf(c).trim()&&!hasPicture(c)&&!all(c,'br').length&&!writingLine){if(!streak)start=indexOf.get(c);streak++;}else flush(start);
      }else if(top&&c.localName==='sectPr')streak=0;
      else{flush(start);if(['tbl','tr','tc','body'].includes(c.localName))walk(c);}
    }
    if(top)streak=0;else flush(start);
  };
  walk(first(doc,'body'),true);
  if(streaks)add('warn','EMPTY-PARAGRAPHS',`${streaks} place${streaks===1?'':'s'} with two or more empty paragraphs in a row.`,streakAt);

  let lowest=Infinity,checked=0,low=false;
  const test=(label,fg,bg)=>{
    const f=hex6(fg)||'000000',b=hex6(bg)||'FFFFFF',r=contrast(f,b);
    checked++;lowest=Math.min(lowest,r);
    if(r<4.5){low=true;add('fail','CONTRAST',`${label}: text ${f} on ${b} is ${r.toFixed(1)}:1; it needs 4.5:1.`);}
  };
  for(const [id,s] of styles){
    if(s.type==='paragraph'){
      const color=resolve(styles,id,x=>hex6(x.color)),fill=resolve(styles,id,x=>hex6(x.fill));
      if(color||fill)test(`Style "${s.name}"`,color,fill);
    }
    if(s.type==='table')for(const sp of kids(s.node,'tblStylePr')){
      const color=val(first(first(sp,'rPr'),'color')),fill=val(first(first(sp,'tcPr'),'shd'),'fill');
      if(color||fill)test(`Table style "${s.name}" (${val(sp,'type')})`,color,fill);
    }
  }
  if(checked&&!low)add('pass','CONTRAST',`${checked} styles with set colors meet 4.5:1 (lowest ${lowest.toFixed(1)}:1).`);

  const counts={fail:0,warn:0,pass:0,info:0};
  for(const f of findings)counts[f.level]++;
  return {findings,counts};
}

// Approximate reading order, to see what a linear reader meets. This is not what any screen reader says.
export function readOutParts({root,stylesRoot}){
  const doc=rootOf(root),styles=readStyles(stylesRoot),lines=[];
  const picture=p=>all(p,'docPr',WP).filter(d=>!all(d,'decorative',ADEC).some(x=>['1','true'].includes(x.getAttribute('val')))).map(d=>`[picture: ${(d.getAttribute('descr')||'no alt text').trim()}]`);
  const paragraph=p=>{
    const id=paragraphStyleId(p),lvl=headingLevel(styles,id),t=textOf(p).trim(),pics=picture(p);
    if(!t&&!pics.length)return null;
    return `${lvl?`Heading ${lvl}`:styleName(styles,id)}: ${[...pics,t].filter(Boolean).join(' ')}`;
  };
  for(const c of els(first(doc,'body'))){
    if(c.localName==='p'){const l=paragraph(c);if(l)lines.push(l);}
    else if(c.localName==='tbl'){
      const rows=kids(c,'tr'),cols=kids(first(c,'tblGrid'),'gridCol').length;
      lines.push(`Table with ${cols} columns and ${rows.length} rows.`);
      rows.forEach((r,i)=>{
        const cells=kids(r,'tc').map(tc=>kids(tc,'p').map(p=>[...picture(p),textOf(p).trim()].filter(Boolean).join(' ')).filter(Boolean).join(' / '));
        lines.push(`  Row ${i+1}${isOn(first(first(r,'trPr'),'tblHeader'))?' (header row)':''}: ${cells.join(' | ')}`);
      });
    }
  }
  return lines;
}
