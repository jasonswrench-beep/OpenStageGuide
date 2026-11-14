// The window carries only the files shell/prepare.mjs lists. If the server loads a script that list leaves out, the installed program
// cannot start. This reads the server's imports the way Node does and checks that every file it reaches is carried.
import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync,existsSync} from 'node:fs';
import {dirname,join,relative,resolve,sep} from 'node:path';
import {fileURLToPath} from 'node:url';

const repo=join(dirname(fileURLToPath(import.meta.url)),'..');
const IMPORT=/(?:^|[\n;])\s*(?:import|export)\s[^'"`]*?from\s*['"]([^'"]+)['"]|(?:^|[\n;\s])import\s*['"]([^'"]+)['"]|import\(\s*['"]([^'"]+)['"]\s*\)/g;

// Every file the server reaches by its imports, and every package (a name that is not ./ or node:) it asks for.
function reached(entry){
  const files=new Set(),packages=new Set(),queue=[resolve(entry)];
  while(queue.length){
    const file=queue.pop();
    if(files.has(file))continue;
    files.add(file);
    const text=readFileSync(file,'utf8');
    for(const m of text.matchAll(IMPORT)){
      const spec=m[1]||m[2]||m[3];
      if(spec.startsWith('node:'))continue;
      if(spec.startsWith('.')){
        const target=resolve(dirname(file),spec);
        assert.ok(existsSync(target),`${relative(repo,file)} imports ${spec}, which is not there`);
        queue.push(target);
      }else packages.add(spec);
    }
  }
  return {files:[...files].map(f=>relative(repo,f)),packages:[...packages]};
}

const prepare=readFileSync(join(repo,'shell','prepare.mjs'),'utf8');
const carriedScripts=[...prepare.match(/for \(const script of \[([^\]]*)\]/)[1].matchAll(/"([^"]+)"/g)].map(m=>m[1]);
// Folders prepare.mjs carries whole.
const carriedFolders=['src','program','public'];
const carriedPackages=[...prepare.matchAll(/copy\(join\("node_modules", "([^"]+)"\)\)/g)].map(m=>m[1]);

test('the setup carries every script and source file the server loads',()=>{
  const {files}=reached(join(repo,'scripts','project-server.mjs'));
  assert.ok(files.length>=5,'the walk found the server and what it loads');
  for(const file of files){
    const [top,...rest]=file.split(sep);
    if(top==='scripts')assert.ok(carriedScripts.includes(rest.join('/')),`${file} is loaded by the server, but shell/prepare.mjs does not carry it`);
    else assert.ok(carriedFolders.includes(top),`${file} is loaded by the server, but shell/prepare.mjs carries nothing of ${top}`);
  }
  assert.ok(carriedScripts.includes('project-server.mjs'),'the server itself is carried');
});

test('the setup carries every package the server asks for, and every script it lists exists',()=>{
  const {packages}=reached(join(repo,'scripts','project-server.mjs'));
  for(const name of packages)assert.ok(carriedPackages.includes(name.split('/')[0]),`the server loads the package ${name}, which shell/prepare.mjs does not carry`);
  for(const script of carriedScripts)assert.ok(existsSync(join(repo,'scripts',script)),`shell/prepare.mjs lists scripts/${script}, which is not there`);
});
