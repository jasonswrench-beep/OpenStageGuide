// Check for updates: whether a newer OpenStage Guide is published. Notify only.
//
// It runs only when the person asks (Help > Check for updates); nothing is checked at start-up, on a timer or in the background. It says
// whether a newer release is published and, if so, gives that release's page on GitHub, where the notes, the checksum and the signed setup are.
// It downloads and installs nothing; the person runs the new setup, which updates the copy in place and keeps the guides.
//
// One HTTPS request, from this server, to GitHub's list of OpenStage Guide's releases (RELEASES_API): the certificate is checked (Node's own),
// no redirect is followed, at most 2 MB is read, and it gives up after 15 seconds. Pre-releases count, because every release so far is one; drafts
// never do. The request carries only a plain user agent; nothing about the person, their guides or their computer is sent.
//
// The release this copy is comes from the window, which the installer's build gives it (--release, from release.txt beside the program, written
// from OPENSTAGE_GUIDE_RELEASE when the window is built). A copy run from the source has none, and is told which release is the newest.
// OPENSTAGE_UPDATE_CHECK=off, --no-update-check, or an update-check.txt saying "off" beside the window's program turns the check off on a
// computer: the answer then says so and nothing is asked of the network. This is the same button, and the same rules, as the other tools of
// the OpenStage family have (docs/UPDATE-CHECK.md).
import {request as httpsRequest} from 'node:https';

/** The repository whose releases are OpenStage Guide's. */
export const REPOSITORY='jasonswrench-beep/OpenStageGuide';
export const RELEASES_PAGE=`https://github.com/${REPOSITORY}/releases`;
export const RELEASES_API=`https://api.github.com/repos/${REPOSITORY}/releases?per_page=30`;
export const USER_AGENT='OpenStage-Guide-update-check';

/** The most of the answer that is read, and how long it is waited for. */
export const MAX_ANSWER_BYTES=2*1024*1024;
export const TIMEOUT_MS=15_000;

const STAGES={alpha:0,beta:1,rc:2};
const VERSION=/^v?(\d{1,6})\.(\d{1,6})\.(\d{1,6})(?:-(alpha|beta|rc)\.(\d{1,6}))?$/;

/**
 * A version as numbers that sort as releases do: 0.1.0-alpha.9 before 0.1.0-alpha.10, the betas after the alphas, and a release with no stage after
 * all of its pre-releases. Anything that is not a plain version is undefined.
 * @param {string} tag
 * @returns {number[] | undefined}
 */
export function versionKey(tag){
  const m=VERSION.exec(String(tag??'').trim());
  if(!m)return undefined;
  const stage=m[4]===undefined?3:STAGES[m[4]];
  return [Number(m[1]),Number(m[2]),Number(m[3]),stage,m[5]===undefined?0:Number(m[5])];
}

/** @param {readonly number[]} a @param {readonly number[]} b */
export function compareVersions(a,b){
  for(let i=0;i<Math.max(a.length,b.length);i++){
    const d=(a[i]??0)-(b[i]??0);
    if(d!==0)return Math.sign(d);
  }
  return 0;
}

/** "0.1.0 alpha 26" for v0.1.0-alpha.26. @param {string} tag */
export function spoken(tag){
  const v=String(tag).trim().replace(/^v/,'');
  const at=v.indexOf('-');
  return at===-1?v:`${v.slice(0,at)} ${v.slice(at+1).replace(/\./g,' ')}`;
}

/** The page of one release of OpenStage Guide, and only of it: a tag that is not a plain version gives the list of releases. @param {string} tag */
export function releasePage(tag){
  return String(tag).startsWith('v')&&versionKey(tag)?`${RELEASES_PAGE}/tag/${tag}`:RELEASES_PAGE;
}

/** Whether an address is a page of OpenStage Guide's releases (the list, or one release by its tag). The window opens no other for this. @param {string} url */
export function isReleasePage(url){
  if(url===RELEASES_PAGE)return true;
  const prefix=`${RELEASES_PAGE}/tag/`;
  const tag=String(url).startsWith(prefix)?String(url).slice(prefix.length):undefined;
  return tag!==undefined&&tag.startsWith('v')&&versionKey(tag)!==undefined;
}

/**
 * The answer for the page, from the text of GitHub's list of releases and the release this copy is.
 * @param {string} text @param {string | undefined} current
 * @returns {{state:'current'|'newer'|'unknown'|'error', message:string, latest?:string, url:string}}
 */
export function answerFrom(text,current){
  let list;
  try{list=JSON.parse(text);}catch{list=undefined;}
  if(!Array.isArray(list))return {state:'error',message:'The list of releases could not be read. Try again later, or look at the releases page.',url:RELEASES_PAGE};
  let newest;
  for(const release of list){
    if(release===null||typeof release!=='object')continue;
    if(release.draft===true||typeof release.tag_name!=='string')continue;
    const tag=release.tag_name;
    const key=tag.startsWith('v')?versionKey(tag):undefined;
    if(key&&(!newest||compareVersions(key,newest.key)>0))newest={key,tag};
  }
  if(!newest)return {state:'error',message:'No release of OpenStage Guide was found. Look at the releases page.',url:RELEASES_PAGE};
  const latest=spoken(newest.tag),url=releasePage(newest.tag);
  const mine=current?versionKey(current):undefined;
  if(!current||!mine)return {state:'unknown',latest,url,message:`The newest release of OpenStage Guide is ${latest}. This copy was not made by the setup, so it cannot say which release it is.`};
  if(compareVersions(newest.key,mine)>0)return {state:'newer',latest,url,message:`OpenStage Guide ${latest} is out. This copy is ${spoken(current)}. Its page has what changed and the setup to download; running the setup updates this copy and keeps your guides.`};
  return {state:'current',latest,url,message:`This copy is the newest release, ${spoken(current)}.`};
}

/**
 * GitHub's list of OpenStage Guide's releases, as text: one GET, the certificate checked, no redirect followed, at most 2 MB, 15 seconds. The address and
 * the request function are parameters only for the tests; the program always asks RELEASES_API.
 * A repository that is not public yet answers 404, which is said as "no release found" (an empty list), not as a failure.
 * @param {string} [url] @param {typeof httpsRequest} [request]
 * @returns {Promise<string>}
 */
export function fetchReleases(url=RELEASES_API,request=httpsRequest){
  return new Promise((resolve,reject)=>{
    const req=request(url,{method:'GET',headers:{'User-Agent':USER_AGENT,Accept:'application/vnd.github+json','X-GitHub-Api-Version':'2022-11-28'},timeout:TIMEOUT_MS},res=>{
      if(res.statusCode===404){res.resume();resolve('[]');return;}
      if(res.statusCode!==200){
        res.resume();
        reject(new Error(res.statusCode>=300&&res.statusCode<400?'GitHub answered with a redirect, which is not followed.':`GitHub answered ${res.statusCode??'nothing'}.`));
        return;
      }
      const parts=[];
      let size=0;
      res.on('data',chunk=>{
        size+=chunk.length;
        if(size>MAX_ANSWER_BYTES){req.destroy(new Error('The list of releases was larger than it should be.'));return;}
        parts.push(chunk);
      });
      res.on('end',()=>resolve(Buffer.concat(parts).toString('utf8')));
      res.on('error',reject);
    });
    req.on('timeout',()=>req.destroy(new Error('GitHub did not answer in time.')));
    req.on('error',reject);
    req.end();
  });
}

/**
 * The answer to Check for updates, with these settings. Never throws: a failure is an answer that says what to do.
 * @param {{current?:string, enabled:boolean, fetchReleases:()=>Promise<string>}} settings
 */
export async function checkForUpdates(settings){
  if(!settings.enabled)return {state:'off',message:'Checking for updates is turned off on this computer.'};
  let text;
  try{text=await settings.fetchReleases();}
  catch{return {state:'error',message:'OpenStage Guide could not reach GitHub to check. Check the internet connection and try again, or look at the releases page.',url:RELEASES_PAGE};}
  return answerFrom(text,settings.current);
}
