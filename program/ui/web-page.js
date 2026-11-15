// @ts-check
/**
 * The guide as one web page (File > Save as web page). It is made from the program's own view of the guide (preview.js), so the page shows what the screen shows,
 * and written as one file: its style sheet and its pictures are inside it. src/guide-html.mjs does the work; this part gets the guide, the program's own icons and
 * gives the file to the author.
 */
import { guideToHtml } from "../../src/guide-html.mjs";
import { previewStyles } from "../../src/guide-styles-edit.mjs";
import { describeGuide } from "./preview.js";
import { downloadFile, safeFileName } from "./word-files.js";

const parseXml = (/** @type {string} */ xml) => new DOMParser().parseFromString(xml, "application/xml");
const ICON_DIR = "/assets/icons-doc/";
const ICON_PATH = /^\/assets\/icons-doc\/[A-Za-z0-9_-]+\.png$/;

/** @param {Uint8Array} bytes */
function toBase64(bytes) {
  let out = "";
  for (let i = 0; i < bytes.length; i += 0x8000) out += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(out);
}

/** The program's own icons the tree names by a path, each fetched once from the program and made a data address. @param {any[]} nodes */
async function iconPictures(nodes) {
  /** @type {Set<string>} */
  const wanted = new Set();
  const walk = (/** @type {any} */ n) => {
    if (n?.tag === "img" && ICON_PATH.test(String(n.attrs?.src ?? ""))) wanted.add(n.attrs.src);
    for (const c of n?.children ?? []) walk(c);
  };
  nodes.forEach(walk);
  /** @type {Map<string, string>} */
  const found = new Map();
  for (const src of wanted) {
    try {
      const res = await fetch(`${ICON_DIR}${src.slice(ICON_DIR.length)}`);
      if (res.ok) found.set(src, `data:image/png;base64,${toBase64(new Uint8Array(await res.arrayBuffer()))}`);
    } catch {
      /* an icon that cannot be fetched is left out: the words beside it say the same */
    }
  }
  return found;
}

/**
 * @param {Document} doc @param {any} guide the guide's document @param {{ name: string, say: (message: string, bad?: boolean) => void }} opts
 */
export async function saveWebPage(doc, guide, { name, say }) {
  try {
    const state = guide.exportState();
    const nodes = describeGuide(state.paragraphs, parseXml, { media: (/** @type {string} */ id) => guide.mediaOf(id), styles: previewStyles(guide.styleEdits(), guide.listStyles()) });
    if (!nodes.length) throw new Error("This guide is empty, so there is no page to make yet.");
    const icons = await iconPictures(nodes);
    const html = guideToHtml({ nodes, title: String(state.title || name), picture: (src) => icons.get(src) ?? null });
    const file = `${safeFileName(name)}.html`;
    downloadFile(doc, new TextEncoder().encode(html), file, "text/html");
    say(`${file} is the guide as one web page. Its pictures are inside it, so it opens in any browser with nothing beside it. It holds the whole guide, including the facilitator's notes and answer keys: for learners, make the workbook.`);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}
