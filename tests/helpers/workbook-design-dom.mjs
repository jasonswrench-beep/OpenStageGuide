// Deterministic DOM contract fixture over real HTML. No browser/native qualification.
const decode=s=>s.replace(/&(?:amp|lt|gt|quot|apos|nbsp|#\d+|#x[\da-f]+);/gi,v=>({'&amp;':'&','&lt;':'<','&gt;':'>','&quot;':'"','&apos;':"'",'&nbsp;':' '}[v]??String.fromCodePoint(v[2].toLowerCase()==='x'?parseInt(v.slice(3,-1),16):Number(v.slice(2,-1)))));
export function makeDOM(html){
 let active=null,root;
 const all=n=>[n,...n.children.flatMap(all)];
 const matches=(n,s)=>{
  if(s.startsWith('#'))return n.id===s.slice(1);
  const attr=/^\[([^=\]]+)(?:=["']?([^"'\]]+)["']?)?\]$/.exec(s);if(attr)return Object.hasOwn(n.attributes,attr[1])&&(attr[2]===undefined||n.getAttribute(attr[1])===attr[2]);
  return n.tagName?.toLowerCase()===s.toLowerCase();
 };
 const query=(scope,selector)=>{const pool=all(scope).slice(1),out=[];for(const part of selector.split(',').map(s=>s.trim())){
  if(part==='main>section'){for(const n of pool)if(n.tagName==='SECTION'&&n.parent?.tagName==='MAIN')out.push(n);continue;}
  const pair=/^(\w+)\s+(\w+)$/.exec(part);if(pair){for(const n of pool)if(matches(n,pair[2]))for(let p=n.parent;p;p=p.parent)if(matches(p,pair[1])){out.push(n);break;}continue;}
  for(const n of pool)if(matches(n,part))out.push(n);
 }return [...new Set(out)];};
 class T{
  constructor(data=''){this.nodeType=3;this.data=String(data);this.parent=null;}
  get textContent(){return this.data;}set textContent(v){this.data=String(v);}
  get parentNode(){return this.parent;}get parentElement(){return this.parent;}
  remove(){if(this.parent)this.parent._childNodes=this.parent._childNodes.filter(n=>n!==this);this.parent=null;}
  cloneNode(){return new T(this.data);}
 }
 class E{
  constructor(tag='div'){this.nodeType=1;this.tagName=tag.toUpperCase();this._childNodes=[];this.attributes={};this.hidden=false;this.disabled=false;this.checked=false;this.rawValue='';this.type=this.tagName==='INPUT'?'text':'';this.name='';this.className='';this.text='';this.validity={};this.parent=null;this.files=[];this.dataset={};this.scrollTop=0;}
  set id(v){this._id=String(v);this.attributes.id=this._id;}get id(){return this._id||'';}
  get parentElement(){return this.parent;}get parentNode(){return this.parent;}
  get children(){return this._childNodes.filter(n=>n.nodeType===1);}
  get childNodes(){return this._childNodes;}get firstChild(){return this._childNodes[0]||null;}get lastChild(){return this._childNodes.at(-1)||null;}
  get text(){return this._childNodes.filter(n=>n.nodeType===3).map(n=>n.textContent).join('');}
  set text(v){for(const n of this._childNodes||[])if(n.nodeType===3)n.parent=null;this._childNodes=(this._childNodes||[]).filter(n=>n.nodeType!==3);if(String(v)){const n=new T(v);n.parent=this;this._childNodes.unshift(n);}}
  get isConnected(){for(let n=this;n;n=n.parent)if(n===root)return true;return false;}
  get options(){return this.tagName==='SELECT'?all(this).filter(n=>n.tagName==='OPTION'):undefined;}
  set value(v){v=String(v);if(this.tagName==='INPUT'&&['text','search','tel','url','email','password'].includes(this.type||'text'))v=v.replace(/[\r\n]/g,'');if(this.tagName==='TEXTAREA')v=v.replace(/\r\n?/g,'\n');this.rawValue=v;this._valueAssigned=true;if(this.validity?.badInput)this.validity={...this.validity,badInput:false};}
  get value(){if(this.tagName==='SELECT'){const options=this.options;if(this.rawValue!==undefined&&options.some(n=>n.value===this.rawValue))return this.rawValue;return this._valueAssigned?'':options[0]?.value||'';}return this.rawValue||'';}
  set textContent(v){this.replaceChildren(...(String(v)?[new T(v)]:[]));}
  get textContent(){return this._childNodes.map(n=>n.textContent).join('');}
  append(...cs){for(let c of cs){if(typeof c==='string')c=new T(c);c.remove();c.parent=this;this._childNodes.push(c);}}
  appendChild(c){this.append(c);return c;}
  insertBefore(c,reference){if(reference===null||reference===undefined){this.append(c);return c;}const index=this._childNodes.indexOf(reference);if(index<0)throw Error('NotFoundError');if(c===reference)return c;c.remove();const at=this._childNodes.indexOf(reference);c.parent=this;this._childNodes.splice(at,0,c);return c;}
  replaceChild(c,old){const index=this._childNodes.indexOf(old);if(index<0)throw Error('NotFoundError');if(c===old)return old;c.remove();const at=this._childNodes.indexOf(old);old.parent=null;c.parent=this;this._childNodes.splice(at,1,c);return old;}
  removeChild(c){if(c.parent!==this)throw Error('NotFoundError');c.remove();return c;}
  replaceChildren(...cs){for(const n of this._childNodes)n.parent=null;this._childNodes=[];this.append(...cs);}
  remove(){if(this.parent)this.parent._childNodes=this.parent._childNodes.filter(n=>n!==this);this.parent=null;}
  replaceWith(...cs){const p=this.parent;if(!p)return;for(const c of cs)p.insertBefore(c,this);this.remove();}
  before(...cs){const p=this.parent;if(p)for(const c of cs)p.insertBefore(c,this);}
  after(...cs){const p=this.parent;if(!p)return;const index=p._childNodes.indexOf(this),ref=p._childNodes[index+1]||null;for(const c of cs)p.insertBefore(c,ref);}
  setAttribute(k,v){v=String(v);this.attributes[k]=v;if(k==='id')this._id=v;if(k==='for')this.htmlFor=v;if(k==='class')this.className=v;if(k==='type')this.type=v;if(k==='name')this.name=v;if(k.startsWith('data-'))this.dataset[k.slice(5).replace(/-([a-z])/g,(_,c)=>c.toUpperCase())]=v;}
  getAttribute(k){return this.attributes[k]??null;}hasAttribute(k){return Object.hasOwn(this.attributes,k);}removeAttribute(k){delete this.attributes[k];if(k==='id')this._id='';}
  getAttributeNames(){return Object.keys(this.attributes);}
  cloneNode(deep=false){const c=new E(this.tagName);for(const[k,v]of Object.entries(this.attributes))c.setAttribute(k,v);for(const k of ['hidden','disabled','checked','rawValue','type','name','className'])c[k]=this[k];c.validity={...this.validity};if(deep)c.append(...this._childNodes.map(n=>n.cloneNode(true)));return c;}
  querySelectorAll(s){return query(this,s);}querySelector(s){return query(this,s)[0]||null;}
  focus(){active=this;}scrollIntoView(){}click(){return this.onclick?.();}addEventListener(k,f){this['on'+k]=f;}setCustomValidity(message){this.validity={...this.validity,customError:!!message};}
  reset(){for(const n of all(this).slice(1))if(['INPUT','TEXTAREA','SELECT'].includes(n.tagName)){n.value=n.getAttribute('value')??(n.tagName==='SELECT'?n.options[0]?.value||'':'');n.checked=n.hasAttribute('checked');}this.resetCount=(this.resetCount||0)+1;}
 }
 root=new E('document');const stack=[root],voids=new Set(['area','base','br','col','embed','hr','img','input','link','meta','param','source','track','wbr']);
 for(const token of html.replace(/<!--[\s\S]*?-->/g,'').matchAll(/<\/?([A-Za-z][\w:-]*)\b([^<>]*?)>|([^<]+)/g)){
  if(token[3]!==undefined){stack.at(-1).append(new T(decode(token[3])));continue;}const tag=token[1].toLowerCase();if(token[0].startsWith('</')){let i=stack.length-1;while(i>0&&stack[i].tagName.toLowerCase()!==tag)i--;if(i>0)stack.length=i;continue;}
  const n=new E(tag);for(const a of token[2].matchAll(/([^\s"'<>\/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+)))?/g)){const k=a[1],v=decode(a[2]??a[3]??a[4]??'');n.setAttribute(k,v);if(k==='value'){n.value=v;n._valueAssigned=true;}if(k==='hidden')n.hidden=true;if(k==='disabled')n.disabled=true;if(k==='checked')n.checked=true;}
  stack.at(-1).append(n);if(!voids.has(tag)&&!token[0].endsWith('/>'))stack.push(n);
 }
 for(const n of all(root))if(n.tagName==='TEXTAREA'&&!n.rawValue)n.value=n.text;
 const document={nodes:null,createElement:tag=>new E(tag),createTextNode:text=>new T(text),getElementById:id=>all(root).find(n=>n.id===id)||null,querySelector:s=>query(root,s)[0]||null,querySelectorAll:s=>query(root,s),get activeElement(){return active;},get focused(){return active;}};
 document.body=all(root).find(n=>n.tagName==='BODY')||root;
 const nodes={get:id=>document.getElementById(id),values:()=>all(root).filter(n=>n.id).values(),has:id=>!!document.getElementById(id)};document.nodes=nodes;
 return {document,root,nodes,routes:query(root,'[data-view]'),all:()=>all(root),n:id=>document.getElementById(id),Element:E};
}
