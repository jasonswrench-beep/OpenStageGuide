// Makes a PDF of the guide, on this computer, for File > Save as PDF. The guide's web page (src/guide-html.mjs) is printed to a PDF by Microsoft Edge, which
// comes with Windows and is the same engine the program's window shows its pages with. Edge writes a tagged PDF when it prints from the command line: the
// headings, paragraphs and tables are marked, and it says its language and title; with --generate-pdf-document-outline it also writes bookmarks from the headings.
// Both are looked for in the file that comes back, not assumed.
//
// The page is written to a temporary folder with a name nobody can guess, Edge is run on it with a fixed list of options (no shell, no address but that file's)
// and a time limit, and the folder is removed again, whatever happens. The page has no script (and says so to the browser: Content-Security-Policy), so
// nothing in a guide can run anything. Nothing is sent anywhere: Edge is not given an address on the network. This does not rest on the page being the program's own:
// the page is read here first, refused if it could reach outside itself, and printed under a policy that this file puts first in its head (see preparePdfPage).
//
// The process that is started is not the one that makes the PDF. Edge as it is now hands the work to a process of its own and the one we started exits (with
// code 0, after a second or two) well before the PDF is there, which took about 13 seconds on a small page. So the file is waited for after the process has
// gone: it counts when it is a whole PDF (it starts with %PDF-, ends with %%EOF) whose size has stopped changing, and the time limit covers all of it. An Edge
// that stays in the foreground and exits when the PDF is written is the same case, found at the first look. The folder holds Edge's profile, which Edge
// keeps open a little longer than the PDF, so it is removed with retries; the guide's page is deleted before that, so it is not left for long.
import {spawn} from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import {existsSync} from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {pathToFileURL} from 'node:url';

export const PDF_MAX_HTML=40*1024*1024;
export const PDF_TIMEOUT_MS=90000;
export const PDF_POLL_MS=250;
/** How long a process that exited with a failure (not code 0) and left no file at all is given before it is called a failure. */
export const PDF_FAIL_GRACE_MS=30000;

/** Where Edge is on this computer (Windows keeps it with the system's program files), or null. */
export function findEdge({env=process.env,exists=existsSync}={}){
  for(const base of [env['ProgramFiles(x86)'],env.ProgramFiles,env.ProgramW6432]){
    if(!base)continue;
    const file=path.join(base,'Microsoft','Edge','Application','msedge.exe');
    if(exists(file))return file;
  }
  return null;
}

/** Whether a PDF has bookmarks: an outline of its headings, which a reader can jump by. @param {Buffer} bytes */
export function pdfHasOutline(bytes){
  return /\/Outlines\s+\d+\s+\d+\s+R/.test(Buffer.from(bytes).toString('latin1'));
}

/** Whether a PDF is marked for screen readers: it holds a structure tree and says it is marked. (The PDF Edge writes keeps these names in the clear.) @param {Buffer} bytes */
export function pdfIsTagged(bytes){
  const text=Buffer.from(bytes).toString('latin1');
  return text.includes('/StructTreeRoot')&&/\/Marked\s+true/.test(text);
}

export class PdfError extends Error{
  constructor(message,code='failed'){super(message);this.name='PdfError';this.code=code;}
}

// ---- The page is read here before it is printed, and printed under a policy of our own.
//
// The page comes from the program's own page (src/guide-html.mjs), which writes only what it allows. This does not rely on that: whatever arrives is read here the way a
// browser reads markup, and refused (not repaired) if it could make Edge reach for anything outside the page: a picture or a style sheet on another computer or on a network
// share, a frame, an object, a redirect, a script. What is let through is the program's own kind of page, whose only addresses are pictures held inside it (data: addresses
// of a png, jpeg or gif) and places in it (#...). And the policy that forbids the rest is put first in the head by us, whatever the page says, so that Edge fetches nothing even
// if something got past this reading.

/** What the page may do when it is printed: fetch nothing, and show only pictures that are inside it. The same words as the guide's own page says (a test holds them together). */
export const PDF_PAGE_POLICY="default-src 'none'; img-src data:; style-src 'unsafe-inline'";
const POLICY_META=`<meta http-equiv="Content-Security-Policy" content="${PDF_PAGE_POLICY}">`;
// Elements a guide's page never has, and that could fetch or run something (the insides of svg and math are read by a browser in a way of their own).
const REFUSED_ELEMENTS=new Set(['script','iframe','frame','frameset','object','embed','applet','link','base','svg','math']);
// Attributes that hold an address: only a picture inside the page and a place in the page may be written in them.
const ADDRESS_ATTRIBUTES=new Set(['src','srcset','href','xlink:href','poster','data','action','formaction','background','ping','manifest','codebase','longdesc','lowsrc','dynsrc']);
const OWN_ADDRESS=/^(?:data:image\/(?:png|jpeg|gif);base64,[A-Za-z0-9+/]+={0,2}|#[A-Za-z0-9_-]*)$/;
// CSS that can name an address or run something. (The page's own style sheet holds none of these, and the styles it writes are held to the same words when it is made.)
const CSS_REACHES=/url\s*\(|image-set|cross-fade|\bimage\s*\(|element\s*\(|@import|expression|javascript|\\/i;
const KNOWN_REFERENCE=/&(?:quot|amp|lt|gt|apos|#34|#39);/g;
const SPACE=/^[\s﻿]*$/;
const DOCTYPE=/<!doctype[\t\n\f\r ]+html[\t\n\f\r ]*>/iy;
const START_NAME=/[A-Za-z][^\t\n\f\r />]*/y;
const ATTRIBUTE_NAME=/[^\t\n\f\r />][^\t\n\f\r />=]*/y;
const UNQUOTED_VALUE=/[^\t\n\f\r >]*/y;
const BLANKS=/[\t\n\f\r ]*/y;
const BLANKS_AND_SLASHES=/[\t\n\f\r /]*/y;

const refuse=what=>{throw new PdfError(`That page ${what}, so it was not printed.`,'refused');};
const plain=text=>String(text).slice(0,24).replace(/[^\x21-\x7e]/g,'?');

/** The characters a style may be written with, once the few character references the page writes are read; any other reference could hide a word from this reading. @param {string} value */
function styleText(value){
  if(/&(?!(?:quot|amp|lt|gt|apos|#34|#39);)/.test(value))refuse('has a style written with a character reference, which a guide\'s page never does');
  return value.replace(KNOWN_REFERENCE,m=>({'&quot;':'"','&lt;':'<','&gt;':'>','&apos;':"'",'&#34;':'"','&#39;':"'",'&amp;':'&'})[m]);
}

/**
 * Reads the page as a browser reads markup (tags, attributes, quotes, the text of a style) and refuses what could reach outside it. Gives where the head begins and ends,
 * for the policy to be put in.
 * @param {string} html
 * @returns {{ doctypeEnd: number, htmlEnd: number, headEnd: number }}
 */
function readPage(html){
  if(html.includes('\u0000'))refuse('holds a null character');
  const found={doctypeEnd:-1,htmlEnd:-1,headEnd:-1};
  const at=(re,pos)=>{re.lastIndex=pos;return re.exec(html);};
  let i=0,lastEnd=0,otherTag=false;
  while((i=html.indexOf('<',i))>=0){
    const next=html[i+1];
    if(next===undefined)break;
    if(next==='!'||next==='?'){
      // Only the opening <!doctype html>, and only first. A comment, a CDATA section or an instruction for a program is read by a browser in ways that differ (and a page of the
      // program has none), so none is let through.
      if(next==='!'&&at(DOCTYPE,i)&&found.doctypeEnd<0&&found.htmlEnd<0&&found.headEnd<0&&!otherTag&&SPACE.test(html.slice(0,i))){found.doctypeEnd=lastEnd=DOCTYPE.lastIndex;i=lastEnd;continue;}
      refuse('holds a comment, a declaration or an instruction for a program, which a guide\'s page never does');
    }
    if(next==='/'){i+=2;continue;}   // an end tag, or text: what it says is not read as a start
    if(!/[A-Za-z]/.test(next)){i+=1;continue;}   // a "<" that is text
    const name=at(START_NAME,i+1)[0].toLowerCase();
    let j=START_NAME.lastIndex;
    const attributes=[];
    for(;;){
      BLANKS_AND_SLASHES.lastIndex=j;BLANKS_AND_SLASHES.exec(html);j=BLANKS_AND_SLASHES.lastIndex;
      if(j>=html.length)refuse('has a tag that is not closed');
      if(html[j]==='>'){j+=1;break;}
      const key=at(ATTRIBUTE_NAME,j)[0].toLowerCase();
      j=ATTRIBUTE_NAME.lastIndex;
      BLANKS.lastIndex=j;BLANKS.exec(html);
      let k=BLANKS.lastIndex,value='';
      if(html[k]==='='){
        k+=1;BLANKS.lastIndex=k;BLANKS.exec(html);k=BLANKS.lastIndex;
        const quote=html[k];
        if(quote==='"'||quote==="'"){
          const end=html.indexOf(quote,k+1);
          if(end<0)refuse('has a tag that is not closed');
          value=html.slice(k+1,end);j=end+1;
        }else{
          UNQUOTED_VALUE.lastIndex=k;UNQUOTED_VALUE.exec(html);
          value=html.slice(k,UNQUOTED_VALUE.lastIndex);j=UNQUOTED_VALUE.lastIndex;
        }
      }
      attributes.push([key,value]);
    }
    if(REFUSED_ELEMENTS.has(name))refuse(`holds a ${plain(name)} element, which a guide's page never does`);
    for(const [key,value] of attributes){
      if(key.startsWith('on')||key==='srcdoc')refuse(`holds a script handler (${plain(key)})`);
      if(ADDRESS_ATTRIBUTES.has(key)&&!OWN_ADDRESS.test(value))refuse(`has an address (${plain(key)}) that points outside the page`);
      if(key==='style'&&CSS_REACHES.test(styleText(value)))refuse('has a style that reaches for something outside the page');
    }
    if(name==='meta'){
      for(const [key,value] of attributes){
        if(key==='http-equiv'&&value.toLowerCase()!=='content-security-policy')refuse(`holds a meta element that acts on the page (${plain(value)})`);
        if(key==='charset'&&value.toLowerCase()!=='utf-8')refuse('is not written in UTF-8');
      }
    }
    if(name==='html'&&found.htmlEnd<0&&found.headEnd<0&&!otherTag){found.htmlEnd=lastEnd=j;}
    else if(name==='head'&&found.headEnd<0){
      // The policy is only read from the head, so the head must come first: nothing but the doctype and the html tag before it.
      if(otherTag||!SPACE.test(html.slice(lastEnd,i)))refuse('is not laid out as the program\'s pages are (something comes before its head)');
      found.headEnd=j;
    }else if(name!=='html'&&name!=='head')otherTag=true;
    if(name==='style'){
      // The text of a style is read as style, up to its end tag.
      const close=/<\/style(?=[\t\n\f\r />])/ig;
      close.lastIndex=j;
      const end=close.exec(html);
      if(!end)refuse('has a style that is not closed');
      if(CSS_REACHES.test(html.slice(j,end.index)))refuse('has a style that reaches for something outside the page');
      i=end.index;
      continue;
    }
    i=j;
  }
  return found;
}

/**
 * The page as it is printed: read and refused (a PdfError, code `refused`) if it could reach outside itself, and with the policy of PDF_PAGE_POLICY as the first thing in its head.
 * @param {string} html
 */
export function preparePdfPage(html){
  const found=readPage(html);
  if(found.headEnd>=0)return html.slice(0,found.headEnd)+POLICY_META+html.slice(found.headEnd);
  const at=found.htmlEnd>=0?found.htmlEnd:found.doctypeEnd>=0?found.doctypeEnd:0;
  return html.slice(0,at)+`<head>${POLICY_META}</head>`+html.slice(at);
}

/** How far a file that Edge is writing has got: not yet a whole PDF ('partial'), not a PDF at all ('bad'), or whole ('complete'). @param {Buffer} bytes */
function pdfProgress(bytes){
  if(bytes.length>=5&&bytes.subarray(0,5).toString('latin1')!=='%PDF-')return 'bad';
  if(bytes.length<5||!bytes.subarray(-32).toString('latin1').replace(/[\s\0]+$/,'').endsWith('%%EOF'))return 'partial';
  return bytes.length<64?'bad':'complete';
}

const sleepMs=ms=>new Promise(resolve=>setTimeout(resolve,ms));

/** Starts Edge (no shell) and watches the process: what it did is in `.exit` ({code,signal}) or `.error`, `sleep` ends early when either happens, `stop` kills it if it still runs. */
function launch(run,edge,args){
  const watch={exit:null,error:null,child:null,wake:null};
  try{watch.child=run(edge,args,{windowsHide:true,stdio:'ignore',shell:false});}catch(e){throw new PdfError(`Edge could not be started: ${e.message}`,'start');}
  watch.child.on('error',e=>{watch.error??=e;watch.wake?.();});
  watch.child.once('exit',(code,signal)=>{watch.exit={code,signal};watch.wake?.();});
  watch.sleep=ms=>new Promise(resolve=>{
    const done=()=>{clearTimeout(timer);watch.wake=null;resolve();};
    const timer=setTimeout(done,ms);watch.wake=done;
  });
  watch.stop=()=>{if(!watch.exit&&!watch.error){try{watch.child.kill();}catch{/* already gone */}}};
  return watch;
}

/**
 * The bytes of the PDF at `out`, once Edge has made it. Nothing is read while the process we started is still running; after it has exited the file is looked at
 * every `pollMs` until it is a whole PDF of the same size twice running, or `deadline` (a time of day, in ms) comes: `timeout`. A file that is not a PDF is `bad-output`
 * at once; no file at all from a process that failed (an exit code other than 0) is `no-output` after `failGraceMs`.
 */
async function waitForPdf(out,watch,{deadline,pollMs,failGraceMs}){
  let lastSize=-1,exitedAt=0;
  for(;;){
    if(watch.error)throw new PdfError(`Edge could not be started: ${watch.error.message}`,'start');
    if(watch.exit){
      exitedAt||=Date.now();
      let bytes=null;
      try{bytes=await fs.readFile(out);}catch{/* not there yet, or Edge still has it open */}
      const seen=bytes?pdfProgress(bytes):'none';
      if(seen==='bad')throw new PdfError('Edge made something that is not a PDF.','bad-output');
      if(seen==='complete'){if(bytes.length===lastSize)return bytes;lastSize=bytes.length;}else lastSize=-1;
      if(seen==='none'&&watch.exit.code!==0&&Date.now()-exitedAt>=failGraceMs)throw new PdfError('Edge did not make a PDF.','no-output');
    }
    const left=deadline-Date.now();
    if(left<=0)throw new PdfError('Edge did not finish making the PDF in time.','timeout');
    await watch.sleep(watch.exit?Math.min(pollMs,left):left);
  }
}

/** Removes the folder. The guide's page goes first and alone; the folder (Edge's profile is in it, and Edge may hold it a little longer) is tried again `tries` times, and once more later if it is still busy. */
async function removeFolder(dir,page,{rm=fs.rm,tries=20,delayMs=250,laterMs=60000}={}){
  const attempt=async target=>{try{await rm(target,{recursive:true,force:true});return true;}catch{return false;}};
  await attempt(page);
  for(let i=1;i<=tries;i++){
    if(await attempt(dir))return;
    if(i<tries)await sleepMs(delayMs);
  }
  setTimeout(()=>{attempt(dir);},laterMs).unref();
}

/**
 * The PDF of a page, as bytes, and whether it is tagged. `run` and `edge` are replaced in tests, and so are the rest (how often the file is looked at, how long a failure is
 * waited for, and the removal of the folder).
 * @param {string} html @param {{ edge?: string | null, run?: typeof spawn, tmp?: string, timeoutMs?: number, pollMs?: number, failGraceMs?: number, cleanup?: { rm?: typeof fs.rm, tries?: number, delayMs?: number, laterMs?: number } }} [options]
 * @returns {Promise<{ bytes: Buffer, tagged: boolean, outlined: boolean }>}
 */
export async function makePdf(html,{edge=findEdge(),run=spawn,tmp=os.tmpdir(),timeoutMs=PDF_TIMEOUT_MS,pollMs=PDF_POLL_MS,failGraceMs=PDF_FAIL_GRACE_MS,cleanup={}}={}){
  if(typeof html!=='string'||!html.trim())throw new PdfError('There is nothing to make a PDF of.','empty');
  if(html.length>PDF_MAX_HTML)throw new PdfError('This guide is too large to make into a PDF here.','too-large');
  if(/<script[\s>]/i.test(html))throw new PdfError('That page holds a script, so it was not printed.','refused');
  const prepared=preparePdfPage(html);
  if(!edge)throw new PdfError('The program makes a PDF with Microsoft Edge, which is not on this computer. Save the Word file and make the PDF in Word instead.','no-edge');
  const dir=await fs.mkdtemp(path.join(tmp,`osg-pdf-${crypto.randomBytes(4).toString('hex')}-`));
  const page=path.join(dir,'page.html'),out=path.join(dir,'out.pdf');
  let watch=null;
  try{
    await fs.writeFile(page,prepared,{encoding:'utf8',flag:'wx'});
    const args=['--headless=new','--disable-gpu','--no-pdf-header-footer','--no-first-run','--no-default-browser-check','--disable-extensions','--disable-sync','--disable-background-networking','--generate-pdf-document-outline',
      `--user-data-dir=${path.join(dir,'profile')}`,`--print-to-pdf=${out}`,pathToFileURL(page).href];
    const deadline=Date.now()+timeoutMs;
    watch=launch(run,edge,args);
    const bytes=await waitForPdf(out,watch,{deadline,pollMs,failGraceMs});
    return {bytes,tagged:pdfIsTagged(bytes),outlined:pdfHasOutline(bytes)};
  }finally{
    watch?.stop();
    await removeFolder(dir,page,cleanup);
  }
}
