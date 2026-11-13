// Writes program/ui/ui-icons.js: the icons the program's menus, buttons and command search draw, as Lucide's own shapes (ISC licence, with
// the Feather notices that go with some of them).
//
//   node scripts/make-ui-icons.mjs --lucide <lucide.json> [--template <a ui-icons.js>] [--out program/ui/ui-icons.js]
//
// --lucide is the file of Lucide icons that OpenStage Author ships (editor/web/icons/lucide.json: {"icons":{name:{s:[shapes]}}}).
// The list of names is program/icons.json. The file's licence header and its uiIcon function are kept from the template: the file already
// written when there is one, or OpenAssess's ui-icons.js (the suite's other plain-JavaScript program) the first time. An unknown icon name
// stops the script: nothing is written.
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const arg=name=>{const i=process.argv.indexOf(name);return i>=0?process.argv[i+1]:undefined;};
const lucidePath=arg('--lucide');
if(!lucidePath){console.error('Name the Lucide file with --lucide.');process.exit(2);}
const out=path.resolve(arg('--out')||path.join(root,'program/ui/ui-icons.js'));
const templatePath=path.resolve(arg('--template')||out);
const list=JSON.parse(fs.readFileSync(path.join(root,'program/icons.json'),'utf8'));
const lucide=JSON.parse(fs.readFileSync(lucidePath,'utf8')).icons;
const names=[...new Set(list.icons)].sort();
const unknown=names.filter(n=>!lucide[n]);
if(unknown.length){console.error(`Not in Lucide: ${unknown.join(', ')}`);process.exit(1);}
const template=fs.readFileSync(templatePath,'utf8');
const marker='export const UI_ICONS = /** @type {const} */ ({';
const start=template.indexOf(marker),end=template.indexOf('\n});\n',start);
if(start<0||end<0){console.error('The template has no UI_ICONS block.');process.exit(1);}
const header=template.slice(0,start).replace("The icons of OpenAssess's own buttons","The icons of the program's own buttons").replace(/^\/\/ Made by .*\n\/\/ .*\n/m,'// Made by scripts/make-ui-icons.mjs from program/icons.json and the Lucide shapes OpenStage Author ships (editor/web/icons/lucide.json):\n// do not edit by hand. Change the list and run the script again.\n');
const tail=template.slice(end+'\n});\n'.length);
const block=names.map(n=>`  ${JSON.stringify(n)}: ${JSON.stringify(lucide[n].s)},`).join('\n');
fs.writeFileSync(out,`${header}${marker}\n${block}\n});\n${tail}`);
console.log(`Wrote ${names.length} icons to ${path.relative(root,out)}`);
