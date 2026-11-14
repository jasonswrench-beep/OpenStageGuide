// Zoom for the document space: how big the guide's page is drawn on the screen, like the zoom of a word processor. It scales the page the guide is drawn on and
// nothing else (the menus, the tools beside it and the dialogs keep their size), and it changes nothing in the guide or in any Word file made from it. The
// level is kept in the program's settings (src/program-settings.mjs) so that it is the same the next time. These are the steps, the limits, the words and the keys,
// in one place for the page and the tests.

export const ZOOM_KEY='osg.zoom.v1';
/** The sizes zoom goes through (1 is the actual size). */
export const ZOOM_STEPS=Object.freeze([0.5,0.67,0.75,0.8,0.9,1,1.1,1.25,1.5,1.75,2,2.5,3]);
export const MIN_ZOOM=ZOOM_STEPS[0];
export const MAX_ZOOM=ZOOM_STEPS[ZOOM_STEPS.length-1];

/** A zoom factor as the program keeps it: a number from MIN_ZOOM to MAX_ZOOM to two places, or 1 for anything else. @param {unknown} value */
export function cleanZoom(value){
  const n=typeof value==='number'||typeof value==='string'&&value.trim()!==''?Number(value):NaN;
  if(!Number.isFinite(n)||n<MIN_ZOOM-1e-9||n>MAX_ZOOM+1e-9)return 1;
  return Math.round(n*100)/100;
}

/** The next size in a direction (1 bigger, -1 smaller) from the one it is now, staying within the limits. @param {unknown} current @param {number} direction */
export function nextZoom(current,direction){
  const now=cleanZoom(current);
  if(direction>0)return ZOOM_STEPS.find(step=>step>now+1e-9)??MAX_ZOOM;
  return [...ZOOM_STEPS].reverse().find(step=>step<now-1e-9)??MIN_ZOOM;
}

/** "Zoom 125%" @param {number} factor */
export const zoomWords=factor=>`Zoom ${Math.round(factor*100)}%`;

/**
 * What a key press asks of zoom, or null: Ctrl+Alt with =, + (bigger), - (smaller) or 0 (actual size). Ctrl+= and Ctrl+Shift+= are subscript and
 * superscript in the guide's words, so zoom has Ctrl+Alt (with Ctrl held, the mouse wheel over the guide does it too). It goes by the character the key
 * makes, so a key that makes another character on someone's keyboard (the Ctrl+Alt of AltGr types characters) is left to type it.
 * @param {{ ctrlKey?: boolean, altKey?: boolean, metaKey?: boolean, key?: string }} event
 * @returns {'more' | 'less' | 'actual' | null}
 */
export function zoomKeyAction(event){
  if(!event.ctrlKey||!event.altKey||event.metaKey)return null;
  if(event.key==='='||event.key==='+')return 'more';
  if(event.key==='-'||event.key==='_')return 'less';
  if(event.key==='0')return 'actual';
  return null;
}
