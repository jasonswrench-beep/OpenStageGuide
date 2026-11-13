import fs from 'node:fs';

// Small deterministic DOM fixture for controller contracts. This is not a
// browser, native keyboard emulator or accessibility qualification.
export class AssessmentElement {
  constructor(tag='div',document){this.tagName=tag.toUpperCase();this.owner=document;this.children=[];this.attributes={};this.hidden=false;this.disabled=false;this.checked=false;this.value='';this.type='';this.name='';this.className='';this.validity={};this.text='';}
  set id(value){this._id=value;if(value&&this.owner)this.owner.nodes.set(value,this);}
  get id(){return this._id||'';}
  set textContent(value){for(const c of this.children)c.parent=null;this.children=[];this.text=String(value);}
  get textContent(){return this.text+this.children.map(c=>c.textContent).join('');}
  append(...nodes){for(const node of nodes){node.remove();node.parent=this;this.children.push(node);}}
  replaceChildren(...nodes){for(const c of this.children)c.parent=null;this.children=[];this.text='';this.append(...nodes);}
  remove(){if(this.parent){this.parent.children=this.parent.children.filter(c=>c!==this);this.parent=null;}}
  setAttribute(name,value){this.attributes[name]=String(value);if(name==='for')this.htmlFor=String(value);}
  getAttribute(name){return this.attributes[name]??null;}
  scrollIntoView(){this.scrolled=true;}
  focus(){this.owner.focused=this;}
}
export function assessmentDom(){
  const html=fs.readFileSync(new URL('../../public/taskpane.html',import.meta.url),'utf-8');
  const document={nodes:new Map(),focused:null,createElement(tag){return new AssessmentElement(tag,this);},getElementById(id){return this.nodes.get(id);}};
  for(const match of html.matchAll(/<([a-z0-9-]+)\b([^>]*\bid="(measure-[^"]+|view-assessments)"[^>]*)>/gu)){
    const node=document.createElement(match[1]);node.id=match[3];node.type=match[2].match(/\btype="([^"]+)"/u)?.[1]||'';
    node.value=match[2].match(/\bvalue="([^"]*)"/u)?.[1]||'';
    node.hidden=/\bhidden(?:\s|$)/u.test(match[2]);
  }
  document.getElementById('view-assessments').hidden=false;
  return {document,html,n:id=>document.getElementById('measure-'+id)};
}
export const descendants=node=>[node,...Array.from(node.children).flatMap(descendants)];
