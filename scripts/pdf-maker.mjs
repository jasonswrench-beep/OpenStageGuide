// Makes a PDF of the guide, on this computer, for File > Save as PDF. The guide's web page (src/guide-html.mjs) is printed to a PDF by Microsoft Edge, which
// comes with Windows and is the same engine the program's window shows its pages with. Edge writes a tagged PDF when it prints from the command line: the
// headings, paragraphs and tables are marked, and it says its language and title; with --generate-pdf-document-outline it also writes bookmarks from the headings.
// Both are looked for in the file that comes back, not assumed.
//
// The page is written to a temporary folder with a name nobody can guess, Edge is run on it with a fixed list of options (no shell, no address but that file's)
// and a time limit, and the folder is removed again, whatever happens. The page has no script (and says so to the browser: Content-Security-Policy), so
// nothing in a guide can run anything. Nothing is sent anywhere: Edge is not given an address on the network.
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
  if(!edge)throw new PdfError('The program makes a PDF with Microsoft Edge, which is not on this computer. Save the Word file and make the PDF in Word instead.','no-edge');
  const dir=await fs.mkdtemp(path.join(tmp,`osg-pdf-${crypto.randomBytes(4).toString('hex')}-`));
  const page=path.join(dir,'page.html'),out=path.join(dir,'out.pdf');
  let watch=null;
  try{
    await fs.writeFile(page,html,{encoding:'utf8',flag:'wx'});
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
