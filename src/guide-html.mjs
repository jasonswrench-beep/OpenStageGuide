// The guide as one web page (the "Save as web page" of the File menu). The page is made from the tree the program's own view of the guide is drawn from
// (program/ui/preview.js: `describeGuide`), so what the page shows is what the guide shows on the screen. It is one file: its style sheet and its pictures are inside it,
// so it opens in any browser with nothing else beside it, and nothing in it reaches out to anywhere.
//
// Nothing from the guide is ever written as markup. A tree holds tags, classes, attributes, words and style properties; each is checked here against a short list of what the
// view uses, and the words and attribute values are escaped. Pure: a tree in, the text of the page out.

const TAGS=new Set(['p','h1','h2','h3','h4','div','span','img','br','table','tr','th','td','blockquote']);
const VOID=new Set(['br','img']);
const ATTRS=new Set(['alt','width','scope','aria-hidden','aria-label']);
const CSS_ALLOWED=new Set(['fontFamily','fontSize','color','fontWeight','fontStyle','textDecoration','textAlign','marginTop','marginBottom','marginLeft','textIndent','lineHeight','backgroundColor','borderTop','borderLeft','borderBottom','borderRight','verticalAlign']);
// A value of a style property is a short run of the characters such values use. Anything that can load, run or break out of the property is refused.
const CSS_VALUE=/^[A-Za-z0-9 #%.,()'"+\-/]{1,120}$/;
const CSS_BAD=/url\s*\(|expression|javascript|@import|<|>|\\/i;
const CLASS=/^[A-Za-z0-9_ -]{1,80}$/;
const DATA_IMAGE=/^data:image\/(?:png|jpeg|gif);base64,[A-Za-z0-9+/]+={0,2}$/;

const esc=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const kebab=k=>k.replace(/[A-Z]/g,c=>`-${c.toLowerCase()}`);
const px=pt=>`${Math.round(pt*133.33)/100}px`;

/** The inline style of a node: its style properties, its paragraph format (as drawFormat draws it) and its height, each only if safe. */
function styleOf(n){
  const parts=[];
  const put=(k,v)=>{const s=String(v);if(CSS_ALLOWED.has(k)&&CSS_VALUE.test(s)&&!CSS_BAD.test(s))parts.push(`${kebab(k)}:${s}`);};
  for(const [k,v] of Object.entries(n.css||{}))put(k,v);
  const f=n.format;
  if(f){
    if(f.alignment)put('textAlign',f.alignment);
    if(f.spaceBefore!==null&&f.spaceBefore!==undefined)put('marginTop',px(f.spaceBefore));
    if(f.spaceAfter!==null&&f.spaceAfter!==undefined)put('marginBottom',px(f.spaceAfter));
    if(f.lineSpacing!==null&&f.lineSpacing!==undefined)put('lineHeight',String(Math.round(f.lineSpacing*120)/100));
    if(f.leftIndent!==null&&f.leftIndent!==undefined)put('marginLeft',px(f.leftIndent));
    if(f.firstLineIndent!==null&&f.firstLineIndent!==undefined)put('textIndent',px(f.firstLineIndent));
  }
  if(Number.isFinite(n.height)&&n.height>0)parts.push(`min-height:${Math.min(2000,Math.round(n.height*1.33))}px`);
  return parts.join(';');
}

/** The words of a heading, as the contents lines name them. */
const words=n=>(n.text??(n.children||[]).map(words).join('')).replace(/\s+/g,' ').trim();
const isHeading=n=>['h1','h2','h3','h4'].includes(n.tag)&&words(n);

/**
 * @param {any[]} nodes the tree from describeGuide
 * @param {{ picture?: (src: string) => string | null, anchors?: Map<string,string> }} [options]
 *   `picture(src)` gives the data address of a picture the tree names by a path (the program's own icons); with none, such a picture is left out.
 */
export function nodesToHtml(nodes,options={}){
  const out=[];
  const anchors=options.anchors||new Map();
  const draw=(n,top=false)=>{
    if(!n||typeof n!=='object')return;
    const tag=TAGS.has(n.tag)?n.tag:'span';
    const attrs=[];
    const cls=n.cls&&CLASS.test(n.cls)?n.cls:'';
    if(cls)attrs.push(['class',cls]);
    for(const [k,v] of Object.entries(n.attrs||{}))if(ATTRS.has(k)&&v!==undefined&&v!==null)attrs.push([k,v]);
    if(tag==='img'){
      const src=String(n.attrs?.src??'');
      const resolved=DATA_IMAGE.test(src)?src:(options.picture?.(src)||null);
      if(!resolved||!DATA_IMAGE.test(resolved))return;
      attrs.push(['src',resolved]);
      if(!attrs.some(([k])=>k==='alt'))attrs.push(['alt','']);
    }
    const anchor=top?anchors.get(n):null;
    if(anchor){if(!/^[a-z0-9-]+$/.test(anchor))throw new Error('A heading was given an address it cannot have.');attrs.push(['id',anchor]);}
    const style=styleOf(n);
    if(style)attrs.push(['style',style]);
    const open=`<${tag}${attrs.map(([k,v])=>` ${k}="${esc(v)}"`).join('')}`;
    if(VOID.has(tag)){out.push(`${open}/>`);return;}
    out.push(open+'>');
    // A contents line is a link to the heading it names, when the guide has one.
    const target=/\bg-toc\b/.test(cls)&&n.text!==undefined?options.anchors?.get(`toc:${words(n)}`):null;
    if(target)out.push(`<a href="#${esc(target)}">${esc(n.text)}</a>`);
    else if(n.text!==undefined)out.push(esc(n.text));
    if(tag==='table'){
      // The rows that are all header cells are the head of the table; the rest are its body.
      const rows=n.children||[],isHead=r=>(r.children||[]).length>0&&r.children.every(c=>c.tag==='th');
      let at=0;
      while(at<rows.length&&isHead(rows[at]))at++;
      if(at){out.push('<thead>');rows.slice(0,at).forEach(r=>draw(r));out.push('</thead>');}
      out.push('<tbody>');rows.slice(at).forEach(r=>draw(r));out.push('</tbody></table>');
      return;
    }
    for(const c of n.children||[])draw(c);
    out.push(`</${tag}>`);
  };
  for(const n of nodes)draw(n,true);
  return out.join('');
}

/** The tree with an address on each heading and each contents line pointed at the heading of the same words. Returns the options `nodesToHtml` takes. */
function withAnchors(nodes,options){
  const anchors=new Map(),byWords=new Map();
  let n=0;
  for(const node of nodes)if(isHeading(node)){
    const id=`s${++n}-${words(node).toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'').slice(0,40)||'part'}`;
    anchors.set(node,id);
    if(!byWords.has(words(node)))byWords.set(words(node),id);
  }
  for(const [w,id] of byWords)anchors.set(`toc:${w}`,id);
  return {...options,anchors};
}

export const WEB_CSS=`
:root{color-scheme:light dark;--text:#1b2433;--muted:#47536a;--bg:#ffffff;--line:#b9c2d3;--tint:#eef2f8;--accent:#0b5c70;--shell:#10243f}
@media (prefers-color-scheme:dark){:root{--text:#e8ecf5;--muted:#b9c2d8;--bg:#0f1b2e;--line:#3a4a68;--tint:#16263f;--accent:#67e8f9;--shell:#e8ecf5}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:1rem/1.55 "Segoe UI",system-ui,-apple-system,sans-serif}
.skip{position:absolute;left:-9999px;top:0;background:var(--accent);color:var(--bg);padding:.5rem 1rem}
.skip:focus{left:0;z-index:10}
main{max-width:46rem;margin:0 auto;padding:1.5rem 1rem 4rem}
a{color:var(--accent)}
h1.g-title{font-size:1.9rem;line-height:1.2;margin:0 0 1.1rem;color:var(--shell)}
h2.g-module,h2.g-heading{font-size:1.35rem;line-height:1.25;margin:2rem 0 .6rem;color:var(--shell)}
h3.g-lesson,h3.g-heading2{font-size:1.1rem;line-height:1.3;margin:1.4rem 0 .4rem;color:var(--shell)}
h4.g-form{font-size:1.05rem;margin:1.5rem 0 .4rem;padding:.4rem .7rem;border-left:5px solid var(--accent);background:var(--tint)}
h4.g-form.facilitator{border-left-color:var(--line)}
p{margin:0 0 .6rem}
.g-time{font-size:.9rem;color:var(--muted);margin:0 0 .7rem}
.g-cue{margin:.6rem 0 .15rem}
.g-cue.facilitator{border-left:4px solid var(--line);padding-left:.7rem;background:var(--tint)}
.g-label{font-weight:700;font-size:.88rem;color:var(--accent);margin-right:.4rem}
.g-icon{width:1.6em;height:1.6em;vertical-align:middle;margin-right:.4rem}
.g-body{margin:0 0 .3rem .8rem}
.g-notes{border-left:4px solid var(--shell);background:var(--tint);padding:.4rem .7rem;margin:.5rem 0;font-size:.93rem}
.g-bullet{position:relative;margin-left:1.5em}
.g-bullet::before{content:"•";position:absolute;left:-1.1em}
.g-toc{margin:0 0 .15rem;font-size:.95rem}
.g-toc-2{margin-left:1.4rem}
.g-formtext,.g-source,.g-print{margin:0 0 .4rem}
.g-source,.g-print{font-size:.88rem;color:var(--muted)}
.g-fact{font-size:.88rem;background:var(--tint);padding:.1rem .5rem;margin:0 0 .1rem}
.g-term{margin:0 0 .5rem}
.g-reference{margin:0 0 .5rem 2.2rem;text-indent:-2.2rem}
.g-caption{font-style:italic;font-size:.88rem;color:var(--muted);margin:.15rem 0 .8rem}
.g-quote{margin:1rem 0 .2rem 1rem;padding-left:.9rem;border-left:4px solid var(--accent);font-style:italic}
.g-quoteby{margin:0 0 1rem 1rem;color:var(--muted)}
.g-cite{font-weight:600}
.g-check{margin-right:.3rem}
.g-rule{border-bottom:1px solid var(--text);height:1.7rem;margin:0 0 .2rem}
.g-box{border:1px solid var(--text);margin:.4rem 0 .8rem;min-height:3.5rem}
.g-field{display:flex;flex-wrap:wrap;gap:.3rem .8rem;align-items:flex-end;margin:.4rem 0}
.g-fieldlabel{font-weight:600}
.g-fieldline{flex:1 1 8rem;border-bottom:1px solid var(--text);min-height:1.4rem}
.g-card{border:1px dashed var(--line);padding:.6rem;margin:.5rem 0}
.g-picture{max-width:100%;height:auto}
.g-picture-missing,.g-other{color:var(--muted)}
.g-blank,.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
table.g-table{border-collapse:collapse;width:100%;margin:.6rem 0 1rem;font-size:.95rem}
table.g-table th,table.g-table td{border:1px solid var(--line);padding:.35rem .55rem;text-align:left;vertical-align:top}
table.g-table th{background:var(--tint)}
@media (forced-colors:active){.g-rule,.g-box,.g-fieldline,.g-card,table.g-table th,table.g-table td{border-color:CanvasText}h4.g-form,.g-quote,.g-notes,.g-cue.facilitator{border-left-color:CanvasText}}
@media print{body{background:#fff;color:#000}.skip{display:none}main{max-width:none;padding:0}a{color:inherit;text-decoration:none}h1,h2,h3,h4{break-after:avoid}table.g-table,.g-card,.g-quote,.g-notes{break-inside:avoid}.g-picture{break-inside:avoid}}
`.trim();

/**
 * The whole page as text. `nodes` is the tree from describeGuide; `title` names the guide; `picture(src)` supplies the data address of the program's own icons.
 * @param {{ nodes: any[], title: string, lang?: string, picture?: (src: string) => string | null }} spec
 */
export function guideToHtml({nodes,title,lang='en',picture}){
  const name=String(title||'Guide').replace(/\s+/g,' ').trim().slice(0,200)||'Guide';
  if(!/^[a-z]{2,3}(-[A-Za-z0-9]{2,8})?$/.test(lang))throw new Error('The language of the page is not one a page can name.');
  const options=withAnchors(nodes,{picture});
  // Each heading carries the address a contents line points at.
  const body=nodes.map(n=>nodesToHtml([n],options)).join('\n');
  return `<!doctype html>
<html lang="${esc(lang)}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'">
<meta name="generator" content="OpenStage Guide">
<title>${esc(name)}</title>
<style>
${WEB_CSS}
</style>
</head>
<body>
<a class="skip" href="#main">Skip to the guide</a>
<main id="main">
${body||'<p class="g-empty">This guide is empty.</p>'}
</main>
</body>
</html>
`;
}
