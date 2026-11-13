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

/** What the Edit menu does with no editor open: nothing, and every command is dimmed. */
const NO_EDIT = { command: () => {}, state: () => ({ canUndo: false, canRedo: false, hasRows: false, inTable: false }) };

/** The Insert menu: each item is a command of `ctx.insert` (see forms.js), after its name. */
const INSERT_ITEMS = /** @type {{ name: string, label: string, icon: UiIconName, title: string }[]} */ ([
  { name: "picture", label: "Picture...", icon: "image", title: "A picture from a file, with a description for people who cannot see it" },
  { name: "form", label: "Form...", icon: "file-text", title: "Start a handout, worksheet, checklist or other form from a template" },
  { name: "lines", label: "Lines to write on...", icon: "pencil-line", title: "Ruled lines for a written answer" },
  { name: "box", label: "Writing box...", icon: "square", title: "An empty box for a drawing or a longer answer" },
  { name: "fields", label: "Row of fields...", icon: "text-cursor-input", title: "Name, Date, Signature: labels with a line to write on after each" },
  { name: "checklist", label: "Check list...", icon: "list-todo", title: "Items with a check box each, and who and when" },
  { name: "table", label: "Table to fill in...", icon: "table", title: "A header row naming each column, and empty rows to write in" },
  { name: "rating", label: "Rating grid...", icon: "grid-3x3", title: "Statements to rate on a scale, with an optional column for evidence" },
  { name: "cards", label: "Cards to cut apart...", icon: "id-card", title: "Cards in a grid with cut lines, to hand out" },
  { name: "source", label: "Source line...", icon: "quote", title: "A short credit or permission line under a form" },
]);

/**
 * The menus, with the commands that do things through `ctx`. `ctx.edit` is the editor's side of the Edit menu: command(name) does one
 * (undo, redo, cut, copy, paste, add, kind, moveUp, moveDown, delete) and state() says what can be done now.
 * `ctx.insert(name)` makes one of the INSERT_ITEMS (a form, lines, a table...) after the row the author is on.
 * @typedef {{ command: (name: string) => unknown, state: () => { canUndo: boolean, canRedo: boolean, hasRows: boolean, inTable: boolean } }} EditApi
 * @param {{ edit?: EditApi, insert?: (name: string) => unknown, hasGuide: () => boolean, openTool: (id: string) => void, updateGuide: () => void, newGuide: () => void, openGuide: () => void, closeGuide: () => void, search: () => void, keys: () => void, about: () => void }} ctx
 * @returns {{ menus: MenuDef[], quick: MenuCommand[] }}  the menus, and the commands of the navy bar while a guide is open
 */
export function buildMenus(ctx) {
  const editor = ctx.edit ?? NO_EDIT;
  const can = (/** @type {"canUndo" | "canRedo" | "hasRows" | "inTable"} */ what) => () => ctx.hasGuide() && editor.state()[what];
  const edits = (/** @type {string} */ name) => () => void editor.command(name);
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
  const edit = {
    id: "edit",
    label: "Edit",
    items: [
      { label: "Undo", icon: "undo-2", keys: "Ctrl+Z", title: "Take back the last change, whether you typed it or a tool made it", run: edits("undo"), enabled: can("canUndo") },
      { label: "Redo", icon: "redo-2", keys: "Ctrl+Y", title: "Bring back what Undo took away", run: edits("redo"), enabled: can("canRedo") },
      "-",
      { label: "Cut", icon: "scissors", keys: "Ctrl+X", title: "Take the paragraphs you are on or have chosen away, to paste elsewhere", run: edits("cut"), enabled: can("hasRows") },
      { label: "Copy", icon: "copy", keys: "Ctrl+C", title: "Copy the paragraphs you are on or have chosen", run: edits("copy"), enabled: can("hasRows") },
      { label: "Paste", icon: "clipboard-paste", keys: "Ctrl+V", title: "Put the paragraphs you cut or copied after the one you are on", run: edits("paste"), enabled: can("hasRows") },
      "-",
      { label: "Add paragraph below", icon: "plus", keys: "Ctrl+Enter", title: "A new paragraph after this one, ready to type in", run: edits("add"), enabled: can("hasRows") },
      { label: "Kind of paragraph...", icon: "pilcrow", title: "Make this paragraph normal text, block text, a facilitator note or a heading", run: edits("kind"), enabled: can("hasRows") },
      { label: "Move up", icon: "arrow-up", keys: "Alt+Up", title: "Move the paragraphs you are on or have chosen above the one before", run: edits("moveUp"), enabled: can("hasRows") },
      { label: "Move down", icon: "arrow-down", keys: "Alt+Down", title: "Move the paragraphs you are on or have chosen below the one after", run: edits("moveDown"), enabled: can("hasRows") },
      { label: "Delete paragraph", icon: "trash", keys: "Delete", title: "Delete the paragraphs you are on or have chosen (Undo brings them back)", run: edits("delete"), enabled: can("hasRows") },
      "-",
      { label: "Add table row below", icon: "plus", keys: "Ctrl+Enter", title: "In a table: an empty row after the one you are in", run: edits("addRow"), enabled: can("inTable") },
      { label: "Delete table row", icon: "trash", keys: "Ctrl+Delete", title: "In a table: delete the row you are in (the header row stays)", run: edits("deleteRow"), enabled: can("inTable") },
    ],
  };
  /** @type {MenuDef} */
  const insert = {
    id: "insert",
    label: "Insert",
    items: INSERT_ITEMS.flatMap((i, n) => [
      ...(n === 2 ? /** @type {MenuEntry[]} */ (["-"]) : []),
      /** @type {MenuCommand} */ ({ label: i.label, icon: i.icon, title: i.title, run: () => void ctx.insert?.(i.name), enabled: ctx.hasGuide }),
    ]),
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
  return { menus: [file, edit, insert, ...middle, help], quick };
}

export function menuEntries(/** @type {MenuDef[]} */ menus, /** @type {string} */ id) {
  return menus.find((m) => m.id === id)?.items ?? [];
}
