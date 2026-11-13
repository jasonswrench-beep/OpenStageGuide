import fs from 'node:fs';
// Minimal DOM contract for the real app module, not a browser or native Word
// substitute. All Office/Word calls terminate in the isolated mock host.
// Which element holds the focus. Like a browser, a disabled element cannot take it.
let focused=null;
export const focusedId=()=>focused?focused.id:null;
export class Element{
  constructor(id=''){this.id=id;this.children=[];this.hidden=false;this.disabled=false;this.value='';this.dataset={};this.attributes={};}
  focus(){if(!this.disabled)focused=this;}
  set value(s){this.rawValue=String(s);if(this.validity?.badInput)this.validity={...this.validity,badInput:false};}
  get value(){return this.rawValue||'';}
  set textContent(s){this.text=String(s);this.children=[];}
  get textContent(){return (this.text||'')+this.children.map(x=>x.textContent).join('');}
  append(...items){for(const item of items){item.remove?.();item.parent=this;this.children.push(item);}}
  replaceChildren(...items){for(const child of this.children)child.parent=null;this.text='';this.children=[];this.append(...items);}
  setAttribute(k,v){this.attributes[k]=v;}
  getAttribute(k){return Object.hasOwn(this.attributes,k)?this.attributes[k]:null;}
  removeAttribute(k){delete this.attributes[k];}
  scrollIntoView(){}
  querySelector(){return this.submitButton||(this.submitButton=new Element());}
  reset(){this.resetCount=(this.resetCount||0)+1;}
  click(){}
  remove(){if(this.parent){this.parent.children=this.parent.children.filter(x=>x!==this);this.parent=null;}}
}
export function dom(){
  focused=null;
  const html=fs.readFileSync(new URL('../../public/taskpane.html',import.meta.url),'utf8'),nodes=new Map([...html.matchAll(/\bid="([^"]+)"/g)].map(([,id])=>[id,new Element(id)]));
  const routes=[...html.matchAll(/data-view="([^"]+)"/g)].map(([,name])=>{const b=new Element();b.dataset.view=name;return b;});
  nodes.get('activity-minutes').value='10';nodes.get('activity-lines').value='5';nodes.get('agenda-start').value='09:00';
  const main=new Element();
  const document={getElementById:id=>nodes.get(id),createElement:()=>new Element(),body:new Element(),
    querySelector:selector=>selector==='main'?main:null,
    querySelectorAll:selector=>selector==='main>section'?[...nodes.values()].filter(x=>x.id.startsWith('view-')):selector==='[data-view]'?routes:selector.startsWith('[data-view],')?[...routes,...nodes.values()]:[]};
  return {nodes,routes,document};
}

export async function settle(){for(let i=0;i<12;i++)await new Promise(resolve=>setImmediate(resolve));}
