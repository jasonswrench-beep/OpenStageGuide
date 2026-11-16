// @ts-check
/**
 * OpenStage Guide as a program: the frame (navy bar, menus), the welcome page, and an open guide (the guide itself, and the tools beside it).
 * The guide's document lives in this page, with its saving (src/guide-project-adapter.mjs). The tools are the pane of the Word add-in, shown in
 * a frame of the page (/program/tools?project=1&program=1); they use the document by message (src/guide-remote.mjs), so there is one document and
 * one writer. The page draws the guide from the document each time it changes.
 */
import { createProjectAdapter } from "../../src/guide-project-adapter.mjs";
import { webProjectStore } from "../../src/guide-project-web.mjs";
import { serveAdapter } from "../../src/guide-remote.mjs";
import { createDemoAdapter, loadIcons } from "../../src/guide-word.mjs";
import { aboutDialog, askFields, askName, chooseGuide, chooseOne } from "./dialogs.js";
import { parseProject, serializeProject } from "../../src/guide-project.mjs";
import { checkForUpdates } from "./updates.js";
import { get, post } from "./api.js";
import { ZOOM_KEY, cleanZoom, nextZoom, zoomWords } from "../../src/zoom.mjs";
import { GUIDES_CLEARED_KEY, cleanClearedAt, splitGuides } from "../../src/program-settings.mjs";
import { startFrame } from "./frame.js";
import { mountEditor } from "./editor.js";
import { askParagraphFormat, askTextFormat } from "./format.js";
import { openStylesPanel } from "./styles-panel.js";
import { INSERTS, newForm } from "./forms.js";
import { drawWelcome } from "./welcome.js";
import { pdfBase64, savePdfFile, saveWebPage } from "./web-page.js";
import { downloadFile, importStylesFromWord, readGuideFromText, readGuideFromWord, safeFileName, saveHandoutPack, saveWordFile } from "./word-files.js";

const doc = document;
const main = /** @type {HTMLElement} */ (doc.getElementById("main"));
const status = /** @type {HTMLElement} */ (doc.getElementById("status"));

/** @type {{ current: string | null, folder: string, guides: { name: string, when: string }[] }} */
let state = { current: null, folder: "", guides: [], version: "", build: "", token: null };
/** @type {HTMLIFrameElement | null} */
let toolsFrame = null;
/** What to do when the tools have loaded: a tool chosen before they were ready. @type {string | null} */
let wantedTool = null;
let toolsReady = false;
/** The open guide's document (with its saving), and the link that lets the tools in the frame use it. @type {any} */
let guideDoc = null;
/** @type {{ stop: () => void, status: (s: { state: string, message?: string }) => void } | null} */
let link = null;
/** @type {number | null} */
let drawTimer = null;
/** The guide as the author writes it, drawn in the page. @type {{ render: () => void, destroy: () => void, focusCursor: () => void, insert: (build: (ctx: any) => string, what: string) => unknown, command: (name: string) => unknown, state: () => { canUndo: boolean, canRedo: boolean, hasRows: boolean, inTable: boolean } } | null} */
let editor = null;

/** @param {string} message @param {boolean} [bad] */
function say(message, bad = false) {
  status.textContent = message;
  status.classList.toggle("error", bad);
  status.hidden = false;
  window.setTimeout(() => {
    if (status.textContent === message) status.textContent = "";
  }, 6000);
}

/** @param {string} iso */
function when(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const days = Math.floor((Date.now() - d.getTime()) / 86400000);
  if (days <= 0) return "Today";
  if (days === 1) return "Yesterday";
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

async function loadState() {
  state = await get("state");
}

/**
 * Zoom (src/zoom.mjs): how big the guide's page is drawn, like the zoom of a word processor. It scales the page only, not the menus, the tools or the
 * dialogs, and it changes nothing in the guide. The level is a style variable on the page (--doc-zoom, which program/style.css gives the guide's page) and is
 * kept in the program's settings, so it is the same the next time.
 */
let zoomNow = 1;
/** @param {number} factor */
const drawZoom = (factor) => {
  doc.documentElement.style.setProperty("--doc-zoom", String(factor));
  doc.dispatchEvent(new CustomEvent("osg-zoom", { detail: { factor } }));
};
async function loadZoom() {
  try {
    zoomNow = cleanZoom((await get("settings"))?.values?.[ZOOM_KEY]);
  } catch {
    zoomNow = 1;
  }
  drawZoom(zoomNow);
}
/** @param {number} factor */
async function setZoom(factor) {
  zoomNow = cleanZoom(factor);
  drawZoom(zoomNow);
  say(zoomWords(zoomNow));
  try {
    await post("settings", { key: ZOOM_KEY, value: zoomNow === 1 ? null : String(zoomNow) });
  } catch {
    /* it is the size on the screen all the same; it is only not kept for next time */
  }
}
const zoom = {
  can: () => state.current !== null,
  more: () => setZoom(nextZoom(zoomNow, 1)),
  less: () => setZoom(nextZoom(zoomNow, -1)),
  actual: () => setZoom(1),
};

/** An Insert menu command: ask what it needs, then put it after the row the author is on. @param {string} name */
async function insertCommand(name) {
  if (!editor || !guideDoc) return;
  try {
    const made = name === "form" ? await newForm(doc, guideDoc) : await INSERTS[name]?.(doc);
    if (made) await editor.insert(made.build, made.what);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/**
 * True once the open guide has been read from its file. A guide that is still opening has nothing in it yet, so a file made from it then would be empty: the
 * command waits for it, and says so when the guide could not be opened at all.
 */
async function guideReady() {
  if (!guideDoc) return false;
  try {
    await guideDoc.ready;
    return true;
  } catch {
    say("The guide could not be opened, so there is nothing to save yet.", true);
    return false;
  }
}

/** The guide as a Word document, or its styles as a Word template. @param {boolean} template */
async function saveWord(template) {
  if (!(await guideReady())) return;
  await saveWordFile(doc, guideDoc, { name: (state.current ?? "guide").replace(/\.guide-project\.json$/i, ""), template, say });
}

/** The guide as one web page. */
async function saveWeb() {
  if (!(await guideReady())) return;
  await saveWebPage(doc, guideDoc, { name: (state.current ?? "guide").replace(/\.guide-project\.json$/i, ""), say });
}

/** The guide as a PDF, made on this computer. */
async function savePdf() {
  if (!(await guideReady())) return;
  await savePdfFile(doc, guideDoc, { name: (state.current ?? "guide").replace(/\.guide-project\.json$/i, ""), say });
}

/** The window's print dialog, with the page set up to print the guide only (program/style.css, "@media print"). What is waiting to be saved is saved first. */
async function printGuide() {
  if (!(await guideReady()) || !(await settle())) return;
  window.print();
}

/** A backup of the guide: a copy of its project file (words, pictures, styles, settings) saved wherever the author chooses. */
async function saveBackup() {
  if (!(await guideReady())) return;
  try {
    const text = serializeProject(guideDoc.exportState());
    const name = (state.current ?? "guide").replace(/\.guide-project\.json$/i, "");
    const file = `${safeFileName(name)} backup ${new Date().toISOString().slice(0, 10)}.guide-project.json`;
    downloadFile(doc, new TextEncoder().encode(text), file, "application/json");
    say(`${file} is a copy of the guide as it is now, with its pictures and settings. Keep one somewhere other than this computer's drive too, such as OneDrive or a USB stick. Restore a backup opens it as a new guide.`);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/** A saved backup read and opened as a new guide of its own, so the guide that is open (or any other) is not changed. */
async function restoreBackup() {
  if (!(await settle())) return;
  try {
    const read = await askFields(doc, {
      title: "Restore a backup",
      intro: "Choose a backup made with Save a backup. It opens as a new guide, so the guide you have open is not changed.",
      confirm: "Read the backup",
      fields: [{ id: "file", label: "Backup file (.guide-project.json)", type: "file", accept: ".json,.guide-project.json" }],
      finish: async (values) => {
        if (!(values.file instanceof File)) throw new Error("Choose a backup file first.");
        return parseProject(await values.file.text());
      },
    });
    if (!read) return;
    const name = await askName(doc, {
      title: "Name of the restored guide",
      label: "Name of the guide",
      confirm: "Restore",
      value: `${read.guide.title || "Restored guide"} (restored)`.slice(0, 80),
      hint: "The backup opens as a new guide with this name. The guide you have open is not changed.",
      accept: (typed) => post("new", { name: typed }),
    });
    if (!name) return;
    // The new guide is empty when it opens; the backup is put into it and saved, and the guide is opened again from the file.
    await show();
    await guideDoc.ready;
    guideDoc.importState(read.guide);
    await guideDoc.flush();
    await show();
    say(`The backup is open as a new guide, "${name}".`);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/**
 * A guide made from a Word file or from pasted lines, opened as a new guide of its own, so the guide that is open (or any other) is not changed.
 * @param {"word" | "text"} kind
 */
async function importGuide(kind) {
  if (!(await settle())) return;
  try {
    const read = kind === "word" ? await readGuideFromWord(doc) : await readGuideFromText(doc);
    if (!read) return;
    const name = await askName(doc, {
      title: "Name of the new guide",
      label: "Name of the guide",
      confirm: "Make the guide",
      value: String(read.title || "Imported guide").slice(0, 80),
      hint: "It opens as a new guide with this name. The guide you have open does not change.",
      accept: (typed) => post("new", { name: typed }),
    });
    if (!name) return;
    // The new guide is empty when it opens; what was read is put into it and saved, and the guide is opened again from the file.
    await show();
    await guideDoc.ready;
    guideDoc.importState(read.guide);
    await guideDoc.flush();
    await show();
    say(`The new guide "${name}" is open. ${read.notes.join(" ")}`);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/** The forms for everyone from a lesson, a module or the whole guide, as one Word file. @param {"lesson" | "module" | "guide"} scope */
async function savePack(scope) {
  if (!(await guideReady())) return;
  await saveHandoutPack(doc, guideDoc, { scope, name: (state.current ?? "guide").replace(/\.guide-project\.json$/i, ""), say });
}

/** Take the styles of a Word file into the guide. */
async function importStyles() {
  if (!guideDoc) return;
  try {
    await importStylesFromWord(doc, guideDoc, say);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

const frame = startFrame(doc, {
  insert: (name) => insertCommand(name),
  saveWord: () => saveWord(false),
  saveBackup: () => saveBackup(),
  restoreBackup: () => restoreBackup(),
  importWord: () => importGuide("word"),
  saveWeb: () => saveWeb(),
  savePdf: () => savePdf(),
  print: () => printGuide(),
  pasteText: () => importGuide("text"),
  saveTemplate: () => saveWord(true),
  importStyles: () => importStyles(),
  zoom,
  edit: {
    command: (name) => editor?.command(name),
    state: () => editor?.state() ?? { canUndo: false, canRedo: false, hasRows: false, inTable: false, onPicture: false },
  },
  guide: () => (state.current ? { name: state.current.replace(/\.guide-project\.json$/i, "") } : null),
  openTool: (id) => openTool(id),
  updateGuide: () => callTools("updateGuide"),
  newGuide: () => void newGuide(),
  openGuide: () => void openGuide(),
  closeGuide: () => void closeGuide(),
  about: () => void aboutDialog(doc, state),
  updates: () => void checkForUpdates(doc),
});

/** @param {string} name */
async function callTools(name, ...args) {
  if (!toolsFrame || !toolsReady) {
    if (name === "openTool") wantedTool = String(args[0]);
    return;
  }
  /** @type {any} */
  const hook = /** @type {any} */ (toolsFrame.contentWindow)?.openstageProgram;
  try {
    await hook?.[name]?.(...args);
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}
/** @param {string} id */
function openTool(id) {
  return callTools("openTool", id);
}

async function showWelcome() {
  closeDocument();
  toolsFrame = null;
  toolsReady = false;
  doc.title = "OpenStage Guide";
  await loadState();
  frame.refresh();
  // The list may have been cleared (a setting of the program): the guides it leaves off are still in the folder, and Open guide lists every one.
  /** @type {string | null} */
  let clearedAt = null;
  try {
    clearedAt = cleanClearedAt((await get("settings"))?.values?.[GUIDES_CLEARED_KEY]);
  } catch {
    clearedAt = null;
  }
  const { shown, hidden } = splitGuides(state.guides, clearedAt);
  drawWelcome(doc, main, {
    guides: shown,
    hidden,
    folder: state.folder,
    when,
    onNew: () => void newGuide(),
    onOpen: () => void openGuide(),
    onOpenNamed: (n) => void openNamed(n),
    onClear: () => void clearList(true),
    onShowAll: () => void clearList(false),
  });
}

/**
 * "Clear previous guides" on the welcome page leaves the guides off its list, and the button that follows brings them back. Only a time is kept (a setting of
 * the program); no guide file is deleted, moved or changed. @param {boolean} clear
 */
async function clearList(clear) {
  try {
    await post("settings", { key: GUIDES_CLEARED_KEY, value: clear ? new Date().toISOString() : null });
    await showWelcome();
    say(clear ? "The list is cleared. No guide was deleted: Open guide still lists every one of them." : "The guides are on the list again.");
    doc.getElementById("recent-title")?.focus();
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/** Draw again soon: a burst of changes is drawn once. */
function scheduleDraw() {
  if (drawTimer !== null) return;
  drawTimer = window.setTimeout(() => {
    drawTimer = null;
    editor?.render();
  }, 30);
}

/** Make the open guide's document, read from its project file, and let go of the one before. */
function openDocument() {
  closeDocument();
  // The cue icons are the real pictures when a Word document is made from the guide; the program's own server has them.
  const made = createProjectAdapter(createDemoAdapter({ icons: (names) => loadIcons(names) }), {
    ...webProjectStore({ token: () => state.token }),
    onStatus: (s) => {
      if (guideDoc !== made) return;
      frame.setSaveState(s.state);
      link?.status(s);
      if (s.state === "error") say(s.message, true);
    },
    onChange: scheduleDraw,
    exportPdf: () => pdfBase64(made, (state.current ?? "guide").replace(/\.guide-project\.json$/i, "")),
  });
  guideDoc = made;
  made.ready.then(
    () => {
      const host = doc.getElementById("work-guide");
      if (guideDoc !== made || !host) return;
      editor = mountEditor(doc, host, made, { zoom: { more: zoom.more, less: zoom.less, actual: zoom.actual, level: () => zoomNow }, say, choose: (options) => chooseOne(doc, options), askFormat: (current, validate) => askParagraphFormat(doc, current, validate), askText: (current, validate) => askTextFormat(doc, current, validate), askName: (options) => askName(doc, options), ask: (options) => askFields(doc, options), openStyles: async (startId) => { if (guideDoc) await openStylesPanel(doc, guideDoc, say, startId); } });
    },
    () => {},
  );
}

function closeDocument() {
  editor?.destroy();
  editor = null;
  link?.stop();
  link = null;
  guideDoc = null;
  if (drawTimer !== null) window.clearTimeout(drawTimer);
  drawTimer = null;
}

/**
 * Save what is waiting before the open guide is left. False when it could not be saved: the guide is then left open, with its changes.
 * A guide that could not be opened has nothing of the author's in it, so it can always be left.
 */
async function settle() {
  if (!guideDoc) return true;
  await guideDoc.flush();
  if (!guideDoc.blocked() && guideDoc.status().state === "error") {
    say("The open guide could not be saved just now, so it was left open with your changes. Try again in a moment.", true);
    return false;
  }
  return true;
}

function showGuide() {
  const name = (state.current ?? "").replace(/\.guide-project\.json$/i, "");
  doc.title = `${name} - OpenStage Guide`;
  toolsReady = false;
  const guide = doc.createElement("section");
  guide.id = "work-guide";
  guide.className = "work-guide";
  guide.setAttribute("aria-label", "The guide");
  const tools = doc.createElement("section");
  tools.className = "work-tools";
  tools.setAttribute("aria-label", "Guide tools");
  const title = doc.createElement("h2");
  title.className = "tools-title";
  title.textContent = "Guide tools";
  toolsFrame = doc.createElement("iframe");
  toolsFrame.id = "tools-frame";
  toolsFrame.title = "Guide tools";
  toolsFrame.src = "/program/tools?project=1&program=1";
  tools.append(title, toolsFrame);
  const work = doc.createElement("div");
  work.className = "work";
  work.append(guide, tools);
  main.replaceChildren(work);
  frame.refresh();
  openDocument();
  // Listening starts now, before the frame has loaded, so its first question is never missed. Only that frame, on this address, is answered.
  link = serveAdapter(guideDoc, {
    post: (m) => toolsFrame?.contentWindow?.postMessage(m, location.origin),
    listen: (fn) => {
      const heard = (/** @type {MessageEvent} */ e) => {
        if (e.origin === location.origin && toolsFrame && e.source === toolsFrame.contentWindow) fn(e.data);
      };
      window.addEventListener("message", heard);
      return () => window.removeEventListener("message", heard);
    },
  });
}

async function show() {
  await loadState();
  if (state.current) {
    frame.setSaveState("");
    showGuide();
  } else await showWelcome();
}

async function newGuide() {
  if (!(await settle())) return;
  // The guide is made while the dialog is open, so a name that cannot be used is refused in the dialog, with what was typed still there.
  const name = await askName(doc, {
    title: "New guide",
    label: "Name of the guide",
    confirm: "Create",
    hint: "For example: Reframing conflict workshop. You can write the guide's title in it later.",
    accept: (typed) => post("new", { name: typed }),
  });
  if (!name) return;
  try {
    await show();
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

async function openGuide() {
  await loadState();
  const name = await chooseGuide(doc, state.guides, when);
  if (name) await openNamed(name);
}

/** @param {string} name */
async function openNamed(name) {
  if (!(await settle())) return;
  try {
    await post("open", { name });
    await show();
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

async function closeGuide() {
  if (!(await settle())) return;
  try {
    await post("close");
    await show();
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

/**
 * The program's window holds the close of the window until this has run (shell/src-tauri/src/main.rs asks for it): the open guide is saved
 * first, then the program is told, and it tells the window, which closes. When the guide cannot be saved the program is told that, the window
 * stays open, and the message that says so is on the page; closing the window again leaves anyway. (Without this, a change made in the last
 * moment before the window went could be lost: saving waits a short while after a change, and a page that is going away cannot be sure its last
 * request is sent.)
 */
Object.defineProperty(window, "osgClosing", {
  value: async () => {
    let saved = true;
    try {
      saved = await settle();
    } catch {
      saved = false;
    }
    try {
      await post("closing", { saved });
    } catch {
      /* the window closes itself after a few seconds */
    }
  },
});

// Save what is waiting when the page is hidden or closed.
window.addEventListener("pagehide", () => void guideDoc?.flush());
doc.addEventListener("visibilitychange", () => {
  if (doc.visibilityState === "hidden") void guideDoc?.flush();
});

// What the tools frame tells the program: it is ready. (Its calls on the document are answered by serveAdapter.)
window.addEventListener("message", (e) => {
  if (e.origin !== location.origin || e.source !== toolsFrame?.contentWindow) return;
  const data = e.data;
  if (!data || typeof data.type !== "string") return;
  if (data.type === "osg-ready") {
    toolsReady = true;
    if (wantedTool) {
      const id = wantedTool;
      wantedTool = null;
      void openTool(id);
    }
  } else if (data.type === "osg-run" && data.command === "form") {
    // The Forms tool's New form button: the same dialog as Insert > Form.
    void insertCommand("form");
  } else if (data.type === "osg-run" && data.command === "backup") {
    void saveBackup();
  } else if (data.type === "osg-run" && data.command === "restore") {
    void restoreBackup();
  } else if (data.type === "osg-run" && data.command === "pack" && ["lesson", "module", "guide"].includes(data.scope)) {
    // The Forms tool's Make handout pack button.
    void savePack(/** @type {"lesson" | "module" | "guide"} */ (data.scope));
  } else if (data.type === "osg-undo" && (data.action === "undo" || data.action === "redo")) {
    // Ctrl+Z or Ctrl+Y pressed in the tools outside a box that is being typed in: the guide's own Undo and Redo, as the Edit menu runs them.
    void editor?.command(data.action);
  } else if (data.type === "osg-zoom" && ["more", "less", "actual"].includes(data.action)) {
    if (zoom.can()) void zoom[/** @type {"more" | "less" | "actual"} */ (data.action)]();
  } else if (data.type === "osg-key" && typeof data.key === "string" && ["k", "K", "/", "p", "P", "F10"].includes(data.key)) {
    // A key the author pressed in the tools that belongs to the program (Ctrl+K, Ctrl+/, Ctrl+P, F10): the program's keys are heard here.
    doc.dispatchEvent(new KeyboardEvent("keydown", { key: data.key, ctrlKey: data.ctrlKey === true, metaKey: data.metaKey === true, bubbles: true, cancelable: true }));
  }
});

void loadZoom();
void show().catch((e) => say(/** @type {Error} */ (e).message, true));
