// @ts-check
/**
 * The menus of the program: every command it has, with the words, icon and keys it has in a menu, in the command search and on a button.
 * The tools come from the one list the add-in's ribbon and pane are built from (src/guide-tools.mjs), so a command cannot exist
 * without its tool or the other way round: each group of the ribbon is a menu here, and each of its twenty-one tools is one item in one menu
 * (Save and Restore a backup are in File, and Help is in Help). Kept apart from the page so it can be tested without one.
 */
import { GROUPS, TOOLS } from "../../src/guide-tools.mjs";

/**
 * @typedef {import("./menubar.js").MenuCommand} MenuCommand
 * @typedef {import("./menubar.js").MenuDef} MenuDef
 * @typedef {import("./menubar.js").MenuEntry} MenuEntry
 * @typedef {import("./ui-icons.js").UiIconName} UiIconName
 */

/** The menu each ribbon group becomes. The Protect group has no menu: its tools go to File and Help. */
const GROUP_MENUS = {
  start: { id: "guide", label: "Guide" },
  write: { id: "write", label: "Write" },
  bring: { id: "bring", label: "Bring in" },
  elearning: { id: "elearning", label: "E-learning" },
  check: { id: "check", label: "Check" },
  make: { id: "make", label: "Make" },
};

/** What each tool is called in a menu. */
export const TOOL_NAMES = /** @type {Record<string, string>} */ ({
  openDesigner: "Guide overview",
  courseSettings: "Course settings",
  courseOutline: "Course outline",
  addModule: "Add module",
  addActivity: "Add activity",
  instructionalCues: "Instructional cues",
  activityLibrary: "Activity library",
  importSlides: "Import slides",
  updateFromSlides: "Update from slides",
  eLearningScreens: "E-learning screens",
  assessments: "Assessments",
  learningAlignment: "Learning alignment",
  reviewGuide: "Review guide",
  timingAndAgenda: "Timing and agenda",
  deliveryVersions: "Delivery versions",
  appearance: "Appearance",
  generateMaterials: "Generate materials",
  participantWorkbook: "Participant workbook",
  saveBackup: "Save a backup...",
  restoreBackup: "Restore a backup...",
  help: "Help",
});

/** @type {Record<string, UiIconName>} */
export const TOOL_ICONS = {
  openDesigner: "layout-dashboard",
  courseSettings: "settings",
  courseOutline: "list-tree",
  addModule: "square-plus",
  addActivity: "square-pen",
  instructionalCues: "message-square",
  activityLibrary: "library",
  importSlides: "presentation",
  updateFromSlides: "refresh-cw",
  eLearningScreens: "monitor-play",
  assessments: "clipboard-list",
  learningAlignment: "target",
  reviewGuide: "list-checks",
  timingAndAgenda: "clock",
  deliveryVersions: "layers",
  appearance: "palette",
  generateMaterials: "file-stack",
  participantWorkbook: "book-open",
  saveBackup: "archive",
  restoreBackup: "archive-restore",
  help: "circle-question-mark",
};

/**
 * The menus, with the commands that do things through `ctx`.
 * @param {{ hasGuide: () => boolean, openTool: (id: string) => void, updateGuide: () => void, newGuide: () => void, openGuide: () => void, closeGuide: () => void, search: () => void, keys: () => void, about: () => void }} ctx
 * @returns {{ menus: MenuDef[], quick: MenuCommand[] }}  the menus, and the commands of the navy bar while a guide is open
 */
export function buildMenus(ctx) {
  /** @param {string} id @returns {MenuCommand} */
  const tool = (id) => {
    const t = /** @type {Record<string, any>} */ (TOOLS)[id];
    return { label: TOOL_NAMES[id], icon: TOOL_ICONS[id], title: t.tip, run: () => ctx.openTool(id), enabled: ctx.hasGuide };
  };
  /** @type {MenuDef} */
  const file = {
    id: "file",
    label: "File",
    items: [
      { label: "New guide...", icon: "file-plus", title: "Start a new guide", run: ctx.newGuide },
      { label: "Open guide...", icon: "folder-open", title: "Open a guide you made before", run: ctx.openGuide },
      { label: "Close guide", icon: "x", title: "Close this guide and go back to the welcome page", run: ctx.closeGuide, enabled: ctx.hasGuide },
      "-",
      tool("saveBackup"),
      tool("restoreBackup"),
    ],
  };
  /** @type {MenuDef} */
  const help = {
    id: "help",
    label: "Help",
    items: [
      tool("help"),
      "-",
      { label: "Search for a command...", icon: "search", keys: "Ctrl+K", title: "Find any command by its name", run: ctx.search },
      { label: "Keyboard keys", icon: "keyboard", keys: "Ctrl+/", title: "Every key the program has", run: ctx.keys },
      "-",
      { label: "About OpenStage Guide", icon: "info", title: "What this program is", run: ctx.about },
    ],
  };
  const middle = GROUPS.filter((g) => /** @type {Record<string, unknown>} */ (GROUP_MENUS)[g.id]).map((g) => {
    const m = /** @type {Record<string, { id: string, label: string }>} */ (GROUP_MENUS)[g.id];
    return /** @type {MenuDef} */ ({ id: m?.id, label: m?.label, items: g.tools.map(tool) });
  });
  /** @type {MenuCommand[]} */
  const quick = [
    { label: "Review guide", icon: "list-checks", title: "Find missing text, objectives, unfinished prompts and timing issues", run: () => ctx.openTool("reviewGuide") },
    { label: "Update guide", icon: "refresh-cw", title: "Renumber the headings and forms, and correct the times", run: ctx.updateGuide },
    { label: "Make workbook", icon: "book-open", title: "Make the participant workbook", run: () => ctx.openTool("participantWorkbook") },
  ];
  return { menus: [file, ...middle, help], quick };
}

export function menuEntries(/** @type {MenuDef[]} */ menus, /** @type {string} */ id) {
  return menus.find((m) => m.id === id)?.items ?? [];
}
