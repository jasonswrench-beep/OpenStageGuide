// @ts-check
/**
 * OpenStage Guide as a program: the frame (navy bar, menus), the welcome page, and an open guide (the guide itself, and the tools beside it).
 * The tools are the pane of the Word add-in, shown in a frame of the page (/program/tools?project=1&program=1), reading and saving the same
 * project through the program's own server. The program draws the guide from that project each time the pane has saved.
 */
import { parseProject } from "../../src/guide-project.mjs";
import { aboutDialog, askName, chooseGuide } from "./dialogs.js";
import { get, post } from "./api.js";
import { startFrame } from "./frame.js";
import { renderGuide } from "./preview.js";
import { drawWelcome } from "./welcome.js";

const doc = document;
const main = /** @type {HTMLElement} */ (doc.getElementById("main"));
const status = /** @type {HTMLElement} */ (doc.getElementById("status"));

/** @type {{ current: string | null, folder: string, guides: { name: string, when: string }[] }} */
let state = { current: null, folder: "", guides: [] };
/** @type {HTMLIFrameElement | null} */
let toolsFrame = null;
/** What to do when the tools have loaded: a tool chosen before they were ready. @type {string | null} */
let wantedTool = null;
let toolsReady = false;

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

const frame = startFrame(doc, {
  guide: () => (state.current ? { name: state.current.replace(/\.guide-project\.json$/i, "") } : null),
  openTool: (id) => openTool(id),
  updateGuide: () => callTools("updateGuide"),
  newGuide: () => void newGuide(),
  openGuide: () => void openGuide(),
  closeGuide: () => void closeGuide(),
  about: () => void aboutDialog(doc),
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
  toolsFrame = null;
  toolsReady = false;
  doc.title = "OpenStage Guide";
  await loadState();
  frame.refresh();
  drawWelcome(doc, main, { guides: state.guides, folder: state.folder, when, onNew: () => void newGuide(), onOpen: () => void openGuide(), onOpenNamed: (n) => void openNamed(n) });
}

async function refreshGuide() {
  const host = doc.getElementById("work-guide");
  if (!host) return;
  try {
    const r = await fetch("/project", { cache: "no-store" });
    const entries = r.status === 204 || !r.ok ? [] : parseProject(await r.text()).guide.paragraphs;
    renderGuide(doc, host, entries, (xml) => new DOMParser().parseFromString(xml, "application/xml"));
  } catch (e) {
    say(`The guide could not be shown: ${/** @type {Error} */ (e).message}`, true);
  }
}

function showGuide() {
  const name = (state.current ?? "").replace(/\.guide-project\.json$/i, "");
  doc.title = `${name} - OpenStage Guide`;
  toolsReady = false;
  const guide = doc.createElement("section");
  guide.id = "work-guide";
  guide.className = "work-guide";
  guide.tabIndex = 0;
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
  void refreshGuide();
}

async function show() {
  await loadState();
  if (state.current) {
    frame.setSaveState("");
    showGuide();
  } else await showWelcome();
}

async function newGuide() {
  const name = await askName(doc, { title: "New guide", label: "Name of the guide", confirm: "Create", hint: "For example: Reframing conflict workshop. You can write the guide's title in it later." });
  if (!name) return;
  try {
    await post("new", { name });
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
  try {
    await post("open", { name });
    await show();
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

async function closeGuide() {
  try {
    await post("close");
    await show();
  } catch (e) {
    say(/** @type {Error} */ (e).message, true);
  }
}

// What the tools frame tells the program: it is ready, how the guide's saving stands, and that a save has finished.
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
  } else if (data.type === "osg-status") {
    frame.setSaveState(String(data.state ?? ""));
    if (data.state === "saved") void refreshGuide();
  }
});

void show().catch((e) => say(/** @type {Error} */ (e).message, true));
