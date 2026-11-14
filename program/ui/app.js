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
import { aboutDialog, askName, chooseGuide, chooseOne } from "./dialogs.js";
import { checkForUpdates } from "./updates.js";
import { get, post } from "./api.js";
import { startFrame } from "./frame.js";
import { mountEditor } from "./editor.js";
import { askParagraphFormat, askTextFormat } from "./format.js";
import { openStylesPanel } from "./styles-panel.js";
import { INSERTS, newForm } from "./forms.js";
import { drawWelcome } from "./welcome.js";
import { importStylesFromWord, saveWordFile } from "./word-files.js";

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

/** The guide as a Word document, or its styles as a Word template. @param {boolean} template */
async function saveWord(template) {
  if (!guideDoc) return;
  await saveWordFile(doc, guideDoc, { name: (state.current ?? "guide").replace(/\.guide-project\.json$/i, ""), template, say });
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
  saveTemplate: () => saveWord(true),
  importStyles: () => importStyles(),
  edit: {
    command: (name) => editor?.command(name),
    state: () => editor?.state() ?? { canUndo: false, canRedo: false, hasRows: false, inTable: false },
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
  drawWelcome(doc, main, { guides: state.guides, folder: state.folder, when, onNew: () => void newGuide(), onOpen: () => void openGuide(), onOpenNamed: (n) => void openNamed(n) });
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
  });
  guideDoc = made;
  made.ready.then(
    () => {
      const host = doc.getElementById("work-guide");
      if (guideDoc !== made || !host) return;
      editor = mountEditor(doc, host, made, { say, choose: (options) => chooseOne(doc, options), askFormat: (current, validate) => askParagraphFormat(doc, current, validate), askText: (current, validate) => askTextFormat(doc, current, validate), askName: (options) => askName(doc, options), openStyles: async (startId) => { if (guideDoc) await openStylesPanel(doc, guideDoc, say, startId); } });
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
  } else if (data.type === "osg-key" && typeof data.key === "string" && ["k", "K", "/", "F10"].includes(data.key)) {
    // A key the author pressed in the tools that belongs to the program (Ctrl+K, Ctrl+/, F10): the program's keys are heard here.
    doc.dispatchEvent(new KeyboardEvent("keydown", { key: data.key, ctrlKey: data.ctrlKey === true, metaKey: data.metaKey === true, bubbles: true, cancelable: true }));
  }
});

void show().catch((e) => say(/** @type {Error} */ (e).message, true));
