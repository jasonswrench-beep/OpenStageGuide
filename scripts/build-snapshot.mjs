import fs from 'node:fs/promises';
import path from 'node:path';
import {createHash} from 'node:crypto';
import {parse} from 'acorn';
const hash=data=>createHash('sha256').update(data).digest('hex');
function freeze(value){Object.freeze(value);for(const child of Object.values(value))if(child&&typeof child==='object')freeze(child);return value;}
const stamp='\n;{const key=Symbol.for("openstage.guide.runtimeModules");(globalThis[key]??=new Set()).add(import.meta.url);}\n';
export async function createBuildSnapshot(root){
  const inputs=new Map(),publicFiles=new Map();
  async function add(relative,url){const full=path.join(root,relative);if((await fs.lstat(full)).isSymbolicLink())throw new Error('Build inputs cannot be symbolic links.');const data=await fs.readFile(full);inputs.set(relative.replaceAll('\\','/'),data);if(url)publicFiles.set(url,data);}
  async function tree(relative,prefix,accept){
    if((await fs.lstat(path.join(root,relative))).isSymbolicLink())throw new Error('Build input folders cannot be symbolic links.');
    for(const e of await fs.readdir(path.join(root,relative),{withFileTypes:true})){if(e.isSymbolicLink())throw new Error('Build inputs cannot be symbolic links.');const name=path.join(relative,e.name),url=prefix+'/'+e.name;if(e.isDirectory())await tree(name,url,accept);else if(accept(e.name))await add(name,url);}
  }
  await tree('src','/src',name=>name.endsWith('.mjs'));
  await tree('public/assets','/assets',name=>/\.(?:png|svg|txt)$/i.test(name));
  for(const name of ['taskpane.html','help.html','picker.html','style.css','vendor/jszip.min.js','vendor/pdf.min.js','vendor/pdf.worker.min.js','vendor/pdfjs-LICENSE.txt'])await add('public/'+name,'/'+name);
  for(const name of ['package.json','package-lock.json','scripts/project-server.mjs','scripts/build-snapshot.mjs','scripts/http-handler.mjs'])await add(name);
  const version=JSON.parse(inputs.get('package.json').toString()).version;
  const descriptor={schema:1,app:'OpenStage Guide',version,inputs:[...inputs].sort(([a],[b])=>a.localeCompare(b,'en')).map(([name,data])=>({path:name,sha256:hash(data),bytes:data.length}))};
  const buildId=hash(JSON.stringify(descriptor)),prefix='/build/'+buildId;
  for(const [url,data] of publicFiles){
    if(url.endsWith('.html'))publicFiles.set(url,Buffer.from(data.toString().replace(/\b(src|href)="\/(?!\/)([^\"]+)"/g,(_,attr,target)=>`${attr}="${prefix}/${target}"`)));
    else if(url.startsWith('/src/')&&url.endsWith('.mjs')){
      const text=data.toString();
      function walk(node){
        if(!node||typeof node!=='object')return;
        if(['ImportDeclaration','ExportNamedDeclaration','ExportAllDeclaration','ImportExpression'].includes(node.type)&&node.source){const spec=node.source.value,target=typeof spec==='string'?path.posix.resolve(path.posix.dirname(url),spec):'';if(typeof spec!=='string'||!spec.startsWith('./')||!target.startsWith('/src/')||!publicFiles.has(target))throw new Error('Runtime imports must use literal relative paths within the pinned source snapshot.');}
        for(const [key,value] of Object.entries(node))if(key!=='source')if(Array.isArray(value))value.forEach(walk);else walk(value);
      }
      walk(parse(text,{ecmaVersion:'latest',sourceType:'module'}));
      publicFiles.set(url,Buffer.from(text+stamp));
    }
  }
  const record={...descriptor,buildId,served:[...publicFiles].sort(([a],[b])=>a.localeCompare(b,'en')).map(([name,data])=>({path:name,sha256:hash(data),bytes:data.length}))};
  freeze(record);return Object.freeze({buildId,version,record,read:url=>publicFiles.has(url)?Buffer.from(publicFiles.get(url)):null});
}
