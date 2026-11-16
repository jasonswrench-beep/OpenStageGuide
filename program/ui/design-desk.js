/**
 * The design desk: where a course is designed before and while the guide is written. Five screens (needs and objectives, models and activities, the course map, the storyboard,
 * the evaluation plan) in front of the guide's own alignment, agenda and timing. The notes are the author's and are kept in the guide's project file (src/design-model.mjs); they
 * are saved a moment after each change and again when the desk is left. The desk takes the place of the guide and its tools while it is open, and "Back to the guide" returns to
 * them as they were. Nothing here needs a model or the network: the Reference Manager's citation is the only thing that talks to another program, and only when asked.
 */
import { el } from "./menubar.js";
import { button } from "./design-ui.js";
import { readDesign, designText } from "../../src/design-model.mjs";
import { needsPanel } from "./design-needs.js";
import { activitiesPanel } from "./design-activities.js";
import { mapPanel } from "./design-mapview.js";
import { boardPanel } from "./design-board.js";
import { evalPanel } from "./design-eval.js";

export const DESK_TABS = Object.freeze([
  { id: "needs", label: "Needs and objectives", make: needsPanel },
  { id: "models", label: "Models and activities", make: activitiesPanel },
  { id: "map", label: "Course map", make: mapPanel },
  { id: "board", label: "Storyboard", make: boardPanel },
  { id: "eval", label: "Evaluation plan", make: evalPanel },
]);
const SAVE_AFTER_MS = 600;

/**
 * Open the desk over the guide. Gives null when the notes in the guide cannot be read (the person is told; nothing is changed).
 * @param {Document} doc
 * @param {{ main: HTMLElement, guide: any, say: (m: string, bad?: boolean) => void, fileName: () => string, title: () => string, appVersion: () => string, newGuideFromPlan: () => Promise<void>, onBack?: () => void }} host
 */
export async function mountDesk(doc, host) {
  let design;
  try {
    design = readDesign(await host.guide.design());
  } catch (e) {
    host.say(/** @type {Error} */ (e).message, true);
    return null;
  }
  const work = host.main.querySelector(".work");
  if (!work) return null;
  const hide = [...work.children].filter((c) => c instanceof HTMLElement && !c.classList.contains("design-desk"));
  /** @type {any} */
  const S = {
    doc,
    guide: host.guide,
    say: host.say,
    design,
    modelId: null,
    tab: "needs",
    fileName: host.fileName,
    title: () => host.title().trim() || host.fileName(),
    appVersion: () => `OpenStage Guide${host.appVersion() ? ` ${host.appVersion()}` : ""}`,
    newGuideFromPlan: host.newGuideFromPlan,
  };

  const desk = el(doc, "section", { class: "design-desk", id: "design-desk", "aria-label": "Design desk" });
  const saveState = el(doc, "p", { class: "dd-save", role: "status", text: "Saved." });
  const back = button(doc, "Back to the guide", () => leave(), { primary: true });
  desk.append(el(doc, "div", { class: "dd-head" }, el(doc, "h2", { text: "Design desk" }), saveState, back));
  const tabs = el(doc, "div", { class: "dd-tabs", role: "tablist", "aria-label": "Design desk screens" });
  const panelHost = el(doc, "div", { class: "dd-panelhost", role: "tabpanel", tabindex: "-1" });
  desk.append(tabs, panelHost);

  // ---- keeping the notes
  let timer = null;
  let dirty = false;
  let saving = Promise.resolve();
  const persist = async () => {
    if (!dirty) return;
    dirty = false;
    try {
      const text = designText(S.design);
      await host.guide.setDesign(text);
      saveState.textContent = "Saved.";
    } catch (e) {
      dirty = true;
      saveState.textContent = "Not saved.";
      host.say(`The design notes could not be saved: ${/** @type {Error} */ (e).message}`, true);
    }
  };
  S.changed = () => {
    dirty = true;
    saveState.textContent = "Saving...";
    if (timer !== null) window.clearTimeout(timer);
    timer = window.setTimeout(() => {
      timer = null;
      saving = saving.then(persist);
    }, SAVE_AFTER_MS);
  };
  const flush = async () => {
    if (timer !== null) {
      window.clearTimeout(timer);
      timer = null;
    }
    saving = saving.then(persist);
    await saving;
  };

  // ---- the screens
  S.render = () => {
    const scroller = desk;
    const top = scroller.scrollTop;
    const tab = DESK_TABS.find((t) => t.id === S.tab) ?? DESK_TABS[0];
    panelHost.setAttribute("aria-labelledby", `dd-tab-${tab.id}`);
    panelHost.replaceChildren(tab.make(S));
    for (const b of tabs.querySelectorAll("button")) {
      const on = b.id === `dd-tab-${tab.id}`;
      b.setAttribute("aria-selected", String(on));
      b.tabIndex = on ? 0 : -1;
    }
    scroller.scrollTop = top;
  };
  S.focus = (prefix) => {
    const want = String(prefix).toLowerCase();
    const label = [...panelHost.querySelectorAll("label")].find((l) => (l.textContent ?? "").trim().toLowerCase().startsWith(want));
    const target = label ? panelHost.querySelector(`#${CSS.escape(label.htmlFor)}`) : [...panelHost.querySelectorAll("button")].find((b) => (b.getAttribute("aria-label") ?? b.textContent ?? "").toLowerCase().startsWith(want));
    /** @type {HTMLElement | null} */ (target)?.focus();
  };
  for (const t of DESK_TABS) {
    const b = el(doc, "button", { type: "button", role: "tab", id: `dd-tab-${t.id}`, "aria-selected": "false", "aria-controls": "dd-panel", text: t.label });
    b.addEventListener("click", () => select(t.id, true));
    b.addEventListener("keydown", (e) => {
      const i = DESK_TABS.findIndex((x) => x.id === t.id);
      const to = e.key === "ArrowRight" ? i + 1 : e.key === "ArrowLeft" ? i - 1 : e.key === "Home" ? 0 : e.key === "End" ? DESK_TABS.length - 1 : -1;
      if (to < 0) return;
      e.preventDefault();
      select(DESK_TABS[(to + DESK_TABS.length) % DESK_TABS.length].id, true);
    });
    tabs.append(b);
  }
  panelHost.id = "dd-panel";
  const select = (id, focus = false) => {
    S.tab = id;
    S.render();
    desk.scrollTop = 0;
    if (focus) /** @type {HTMLElement | null} */ (tabs.querySelector(`#dd-tab-${id}`))?.focus();
  };

  const leave = async () => {
    await flush();
    hide.forEach((c) => c.removeAttribute("hidden"));
    desk.remove();
    host.onBack?.();
  };

  work.append(desk);
  hide.forEach((c) => c.setAttribute("hidden", ""));
  select("needs");
  return {
    /** Show one screen (and bring the desk forward). @param {string} id */
    show(id) {
      if (!desk.isConnected) {
        work.append(desk);
        hide.forEach((c) => c.setAttribute("hidden", ""));
      }
      select(DESK_TABS.some((t) => t.id === id) ? id : S.tab, true);
    },
    isOpen: () => desk.isConnected,
    /** Save what is waiting. */
    flush,
    /** Leave the desk. */
    leave,
  };
}
