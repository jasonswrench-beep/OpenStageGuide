import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const css=fs.readFileSync(new URL('../public/style.css',import.meta.url),'utf8');
const html=fs.readFileSync(new URL('../public/taskpane.html',import.meta.url),'utf8');

// Parse this stylesheet's flat rules and media blocks without starting a browser.
// This covers declared colors/cascade, not native rendering or forced-color paint.
function rules(source,media=[]){
  const result=[];source=source.replace(/\/\*[\s\S]*?\*\//g,'');let pos=0;
  while(pos<source.length){
    const open=source.indexOf('{',pos);if(open<0)break;
    const selector=source.slice(pos,open).trim();let depth=1,end=open+1,quote='';
    for(;end<source.length&&depth;end++){
      const c=source[end];
      if(quote){if(c===quote&&source[end-1]!=='\\')quote='';continue;}
      if(c==='"'||c==="'"){quote=c;continue;}
      if(c==='{')depth++;else if(c==='}')depth--;
    }
    assert.equal(depth,0,'Stylesheet block closes');
    const body=source.slice(open+1,end-1);pos=end;
    if(selector.startsWith('@media'))result.push(...rules(body,[...media,selector]));
    else{
      const declarations=Object.fromEntries(body.split(';').filter(x=>x.trim()).map(x=>{
        const colon=x.indexOf(':');assert.ok(colon>0);return [x.slice(0,colon).trim(),x.slice(colon+1).trim()];
      }));
      result.push({selectors:selector.split(',').map(x=>x.trim()),declarations,media});
    }
  }
  return result;
}
const sheet=rules(css);
const declaration=(selector,name)=>{
  const matches=sheet.filter(r=>!r.media.length&&r.selectors.includes(selector)&&Object.hasOwn(r.declarations,name));
  assert.ok(matches.length,`${selector} declares ${name}`);return matches.at(-1).declarations[name];
};
const resolve=value=>value.replace(/var\((--[\w-]+)\)/g,(_,name)=>declaration(':root',name));
function rgba(value){
  let hex=resolve(value).toLowerCase();assert.match(hex,/^#[a-f0-9]{3,8}$/);
  hex=hex.slice(1);if(hex.length===3||hex.length===4)hex=[...hex].map(c=>c+c).join('');
  assert.ok(hex.length===6||hex.length===8);
  return [0,2,4].map(i=>parseInt(hex.slice(i,i+2),16)).concat(hex.length===8?parseInt(hex.slice(6),16)/255:1);
}
const composite=(foreground,background)=>foreground.slice(0,3).map((v,i)=>v*foreground[3]+background[i]*(1-foreground[3]));
const luminance=rgb=>rgb.slice(0,3).map(v=>v/255).map(v=>v<=0.04045?v/12.92:((v+0.055)/1.055)**2.4).reduce((sum,v,i)=>sum+v*[0.2126,0.7152,0.0722][i],0);
const contrast=(a,b)=>{const x=luminance(a),y=luminance(b);return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05);};
function focusFor(inHeader){
  const matches=sheet.flatMap((r,order)=>r.media.length?[]:r.selectors.filter(s=>['button','button:focus-visible',...(inHeader?['.build-info button','.build-info button:focus-visible']:[])].includes(s)).map(s=>({declarations:r.declarations,order,specificity:(s.match(/[.:]/g)||[]).length*10+1})));
  const focus={};
  for(const {declarations:d} of matches.sort((a,b)=>a.specificity-b.specificity||a.order-b.order)){
    if(d.outline){const m=d.outline.match(/^(\d+(?:\.\d+)?)px\s+(\w+)\s+(#[a-f\d]+)$/i);assert.ok(m,'Outline remains explicit');Object.assign(focus,{width:Number(m[1]),style:m[2],color:rgba(m[3])});}
    if(d['outline-color'])focus.color=rgba(d['outline-color']);
    if(d['outline-offset']){assert.match(d['outline-offset'],/^\d+(?:\.\d+)?px$/);focus.offset=Number.parseFloat(d['outline-offset']);}
  }
  return focus;
}

test('the header build controls stay unique, visible and separate from support details',()=>{
  const header=html.match(/<header\b[^>]*>([\s\S]*?)<\/header>/)?.[1];
  const support=html.match(/<details\b[^>]*class="support"[^>]*>([\s\S]*?)<\/details>/)?.[1];
  assert.ok(header);assert.ok(support);
  for(const id of ['build-loaded','build-check','build-download','build-diagnostic-summary'])assert.equal([...html.matchAll(new RegExp(`\\bid="${id}"`,'g'))].length,1,`${id} is unique`);
  assert.match(header,/<p\b[^>]*id="build-loaded"[^>]*role="status"/);
  assert.match(header,/<button\b[^>]*id="build-check"[^>]*type="button"[^>]*>Check loaded build<\/button>/);
  assert.doesNotMatch(header,/\bhidden\b|<details\b/);
  assert.doesNotMatch(support,/id="build-loaded"|id="build-check"/);
  assert.match(support,/id="build-download"/);assert.match(support,/id="build-diagnostic-summary"/);
  assert.ok(html.indexOf('id="build-check"')<html.indexOf('<nav'));
});

test('declared header focus contrast passes while the inherited amber regression fails',()=>{
  const headerBackground=rgba(declaration('header','background'));
  const panelBackground=composite(rgba(declaration('.build-info','background')),headerBackground);
  const focus=focusFor(true),globalFocus=focusFor(false);
  assert.ok(focus.width>=2);assert.equal(focus.style,'solid');assert.ok(focus.offset>=2);
  assert.ok(contrast(focus.color,panelBackground)>=3,'Header focus contrasts with its actual composed adjacent background');
  assert.ok(contrast(globalFocus.color,panelBackground)<3,'The former inherited color is a meaningful failing control');
  assert.deepEqual(globalFocus.color,rgba('#b56b18'),'Global main-content focus color is preserved');
  assert.ok(contrast(globalFocus.color,rgba(declaration(':root','background')))>=3,'Main-content focus retains contrast');
  assert.equal(globalFocus.width,3);assert.equal(globalFocus.offset,3);
  assert.equal(sheet.filter(r=>!r.media.length&&r.selectors.includes('.build-info button:focus-visible')).length,1);
});

test('header status text and forced-color adaptation remain intact',()=>{
  const background=composite(rgba(declaration('.build-info','background')),rgba(declaration('header','background')));
  for(const selector of ['.build-info p','.build-info p[data-status=stale]','.build-info p[data-status=mixed]'])assert.ok(contrast(rgba(declaration(selector,'color')),background)>=4.5,`${selector} retains normal-text contrast`);
  assert.equal(declaration('.build-info p','overflow-wrap'),'anywhere');
  assert.doesNotMatch(css,/forced-color-adjust\s*:\s*none/i);
  const forced=sheet.filter(r=>r.media.some(m=>/^@media\s*\(forced-colors\s*:\s*active\)/.test(m)));
  assert.ok(forced.some(r=>r.selectors.includes('button')&&r.declarations.border==='1px solid ButtonText'),'System-color border support is retained');
});
