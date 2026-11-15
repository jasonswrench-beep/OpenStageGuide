// Draws the five cue pictures that were added with the script-mechanics cues (docs/KIT-SPEC-2026-11-15.md, K3) as simple stand-ins in the style of the others
// (a black ring, a bold black picture): the vector file goes to docs/icons-source and a 192-pixel PNG to public/assets/icons-doc. The author's own artwork
// replaces any of them by running scripts/make-cue-icons.ps1 on the revised icon. It needs Playwright, which this repository does not depend on:
//   set PLAYWRIGHT_DIR to the playwright package folder, then run  node scripts/make-stand-in-icons.mjs  from the repository root.
import {createRequire} from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
const require=createRequire(import.meta.url);
if(!process.env.PLAYWRIGHT_DIR){console.error('Set PLAYWRIGHT_DIR to the playwright package folder.');process.exit(2);}
const {chromium}=require(process.env.PLAYWRIGHT_DIR);
const ROOT=path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/,'$1')),'..');
const RING='<circle cx="96" cy="96" r="88" fill="#fff" stroke="#000" stroke-width="8"/>';
const ICONS={
  // a speech bubble with an A in it: the answer to listen for
  'expected-answer':'<path d="M40 52h112a14 14 0 0 1 14 14v52a14 14 0 0 1-14 14h-40l-22 22v-22H40a14 14 0 0 1-14-14V66a14 14 0 0 1 14-14z" transform="translate(-2 -6)" fill="#000"/><path d="M94 60 120 112h-12l-5-11H85l-5 11H68L94 60zm0 18-7 14h14z" transform="translate(2 -2)" fill="#fff" fill-rule="evenodd"/>',
  // two bubbles that overlap: talking it over afterwards
  'debrief':'<path d="M34 50h64a10 10 0 0 1 10 10v36a10 10 0 0 1-10 10H70l-16 14V106H34a10 10 0 0 1-10-10V60a10 10 0 0 1 10-10z" fill="#000"/><path d="M86 82h66a10 10 0 0 1 10 10v36a10 10 0 0 1-10 10h-14v14l-16-14H86a10 10 0 0 1-10-10V92a10 10 0 0 1 10-10z" fill="#fff" stroke="#000" stroke-width="8"/>',
  // two arrows going opposite ways: this instead of that
  'option':'<path d="M44 72h76" stroke="#000" stroke-width="12" stroke-linecap="round" fill="none"/><path d="M110 50l26 22-26 22z" fill="#000"/><path d="M148 122H72" stroke="#000" stroke-width="12" stroke-linecap="round" fill="none"/><path d="M82 100l-26 22 26 22z" fill="#000"/>',
  // a table with a chair to each side: the room made ready
  'set-up':'<rect x="52" y="76" width="88" height="14" rx="4" fill="#000"/><rect x="62" y="90" width="12" height="40" fill="#000"/><rect x="118" y="90" width="12" height="40" fill="#000"/><rect x="30" y="94" width="14" height="36" rx="3" fill="#000"/><rect x="148" y="94" width="14" height="36" rx="3" fill="#000"/>',
  // a broom: putting it all away
  'clean-up':'<path d="M122 36 100 106" stroke="#000" stroke-width="12" stroke-linecap="round" fill="none"/><path d="M78 104h44l16 46H62z" fill="#000"/><path d="M84 124l-4 22M96 124v22M108 124l4 22" stroke="#fff" stroke-width="4" stroke-linecap="round" fill="none"/>'
};
fs.mkdirSync(path.join(ROOT,'docs/icons-source'),{recursive:true});
const browser=await chromium.launch();
try{
  const page=await browser.newPage({viewport:{width:192,height:192},deviceScaleFactor:1});
  for(const [id,glyph] of Object.entries(ICONS)){
    const svg=`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 192 192" width="192" height="192">${RING}${glyph}</svg>`;
    fs.writeFileSync(path.join(ROOT,`docs/icons-source/${id}.svg`),svg+'\n');
    await page.setContent(`<html><body style="margin:0;background:transparent">${svg}</body></html>`);
    await page.screenshot({path:path.join(ROOT,`public/assets/icons-doc/${id}.png`),omitBackground:true,clip:{x:0,y:0,width:192,height:192}});
    console.log('drawn',id);
  }
}finally{await browser.close();}
