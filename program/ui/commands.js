// @ts-check
/**
 * The menus of the program: every command it has, with the words, icon and keys it has in a menu, in the command search and on a button.
 * The tools come from the one list the add-in's ribbon and pane are built from (src/guide-tools.mjs), so a command cannot exist
 * without its tool or the other way round: each of its twenty-one tools and each of the four the program adds (PROGRAM_TOOLS) is one item in one menu.
 * The menus are the program's own (PROGRAM_MENUS: Start, Sections, Activities, Slides, Assessment, eLearning, Check & Plan, Make), grouped by the work
 * they are for and not in threes, which only the ribbon needed. Save and Restore a backup are in File, and Help is in Help. Kept apart from the page so
 * it can be tested without one.
 */
import { PROGRAM_MENUS, PROGRAM_NAMES, PROGRAM_TOOLS, TOOLS } from "../../src/guide-tools.mjs";

/**
 * @typedef {import("./menubar.js").MenuCommand} MenuCommand
 * @typedef {import("./menubar.js").MenuDef} MenuDef
 * @typedef {import("./menubar.js").MenuEntry} MenuEntry
 * @typedef {import("./ui-icons.js").UiIconName} UiIconName
 */

/**
 * What each tool is called in a menu: its program name (src/guide-tools.mjs, the same name "More tools" and the tool's heading use), and the two that
 * ask for a file get the three dots a command with a dialog has.
 */
export const TOOL_NAMES = /** @type {Record<string, string>} */ ({
  ...PROGRAM_NAMES,
  saveBackup: `${PROGRAM_NAMES.saveBackup}...`,
  restoreBackup: `${PROGRAM_NAMES.restoreBackup}...`,
});

/** @type {Record<string, UiIconName>} */
export const TOOL_ICONS = {
  openDesigner: "layout-dashboard",
  courseSettings: "settings",
  courseOutline: "list-tree",
  addModule: "square-plus",
  addLesson: "list-plus",
  learningOutcomes: "goal",
  addActivity: "square-pen",
  instructionalCues: "message-square",
  activityLibrary: "library",
  forms: "file-text",
  importSlides: "presentation",
  updateFromSlides: "refresh-cw",
  eLearningScreens: "monitor-play",
  eLearningCues: "mouse-pointer-click",
  accessibility: "accessibility",
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

/** The Format menu's items, each a command of the editor (see editor.js). */
const FORMAT_ITEMS = /** @type {({ name: string, label: string, icon: UiIconName, keys?: string, title: string, own?: boolean } | "-")[]} */ ([
  { name: "bold", label: "Bold", icon: "bold", keys: "Ctrl+B", title: "Make the chosen words bold, or take it off" },
  { name: "italic", label: "Italic", icon: "italic", keys: "Ctrl+I", title: "Slant the chosen words, or take it off" },
  { name: "underline", label: "Underline", icon: "underline", keys: "Ctrl+U", title: "Underline the chosen words (underlined words look like links on a screen, so bold or italic is clearer)" },
  { name: "strike", label: "Strikethrough", icon: "strikethrough", title: "Draw a line through the chosen words" },
  { name: "superscript", label: "Superscript", icon: "superscript", keys: "Ctrl+Shift+=", title: "Raise the chosen words above the line, as in 10 to the power of 3" },
  { name: "subscript", label: "Subscript", icon: "subscript", keys: "Ctrl+=", title: "Lower the chosen words below the line, as in H2O" },
  { name: "text", label: "Text...", icon: "type", title: "The font, size and color of the chosen words, and their looks, in one place" },
  { name: "charStyle", label: "Character style...", icon: "baseline", title: "Give the chosen words one of the character styles you made, or take it off" },
  { name: "clearText", label: "Clear text formatting", icon: "remove-formatting", keys: "Ctrl+Space", title: "Take away the looks given to the chosen words, so the paragraph's style decides again" },
  "-",
  { name: "alignLeft", label: "Align left", icon: "text-align-start", keys: "Ctrl+L", title: "Line the paragraphs up on the left, the easiest to read" },
  { name: "alignCenter", label: "Center", icon: "text-align-center", keys: "Ctrl+E", title: "Put the paragraphs in the middle of the line" },
  { name: "alignRight", label: "Align right", icon: "text-align-end", keys: "Ctrl+R", title: "Line the paragraphs up on the right" },
  { name: "alignJustify", label: "Justify", icon: "text-align-justify", keys: "Ctrl+J", title: "Spread the lines to both edges (the gaps between words vary, which is harder to read)" },
  "-",
  { name: "paragraph", label: "Paragraph...", icon: "pilcrow", title: "Alignment, the space around a paragraph, line spacing and indents" },
  { name: "clearFormat", label: "Clear paragraph formatting", icon: "eraser", title: "Let the paragraphs' styles decide how they look again" },
  "-",
  { name: "styles", label: "Styles...", icon: "palette", title: "Change how a kind of paragraph looks, for every paragraph in it, or make a style of your own" },
  { name: "newStyle", label: "New style from this paragraph...", icon: "square-plus", title: "Keep this paragraph's alignment and spacing as a style, and put the paragraph in it" },
  { name: "importStyles", label: "Styles from a Word file...", icon: "file-input", title: "Take the styles of a Word document or template into the guide", own: true },
]);

/**
 * The menus, with the commands that do things through `ctx`. `ctx.edit` is the editor's side of the Edit menu: command(name) does one
 * (undo, redo, cut, copy, paste, add, kind, moveUp, moveDown, delete) and state() says what can be done now.
 * `ctx.insert(name)` makes one of the INSERT_ITEMS (a form, lines, a table...) after the row the author is on.
 * @typedef {{ command: (name: string) => unknown, state: () => { canUndo: boolean, canRedo: boolean, hasRows: boolean, inTable: boolean } }} EditApi
 * @param {{ edit?: EditApi, insert?: (name: string) => unknown, saveWord?: () => unknown, saveTemplate?: () => unknown, saveBackup?: () => unknown, restoreBackup?: () => unknown, importStyles?: () => unknown, zoom?: { more: () => unknown, less: () => unknown, actual: () => unknown, can: () => boolean }, hasGuide: () => boolean, openTool: (id: string) => void, updateGuide: () => void, newGuide: () => void, openGuide: () => void, closeGuide: () => void, search: () => void, keys: () => void, about: () => void, updates?: () => void }} ctx
 * @returns {{ menus: MenuDef[], quick: MenuCommand[] }}  the menus, and the commands of the navy bar while a guide is open
 */
export function buildMenus(ctx) {
  const editor = ctx.edit ?? NO_EDIT;
  const can = (/** @type {"canUndo" | "canRedo" | "hasRows" | "inTable"} */ what) => () => ctx.hasGuide() && editor.state()[what];
  const edits = (/** @type {string} */ name) => () => void editor.command(name);
  /** @param {string} id @returns {MenuCommand} */
  const tool = (id) => {
    const t = /** @type {Record<string, any>} */ (TOOLS)[id] ?? /** @type {Record<string, any>} */ (PROGRAM_TOOLS)[id];
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
      { label: "Save as Word file...", icon: "file-down", title: "Make a Word document from the guide, with its styles, pictures and a contents table. Word can save it as a PDF", run: () => void ctx.saveWord?.(), enabled: ctx.hasGuide },
      { label: "Save styles as a Word template...", icon: "swatch-book", title: "Make a Word template that has the guide's styles, so other documents can look like the guide", run: () => void ctx.saveTemplate?.(), enabled: ctx.hasGuide },
      "-",
      { label: TOOL_NAMES.saveBackup, icon: TOOL_ICONS.saveBackup, title: "Save a copy of the guide as it is now, with its pictures and settings, wherever you choose", run: () => void ctx.saveBackup?.(), enabled: ctx.hasGuide },
      { label: TOOL_NAMES.restoreBackup, icon: TOOL_ICONS.restoreBackup, title: "Open a backup you saved earlier as a new guide; the guide you have open does not change", run: () => void ctx.restoreBackup?.() },
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
  const format = {
    id: "format",
    label: "Format",
    items: FORMAT_ITEMS.map((i) =>
      i === "-"
        ? i
        : /** @type {MenuCommand} */ (i.own ? { label: i.label, icon: i.icon, title: i.title, run: () => void ctx.importStyles?.(), enabled: ctx.hasGuide } : { label: i.label, icon: i.icon, keys: i.keys, title: i.title, run: edits(i.name), enabled: can("hasRows") }),
    ),
  };
  // The View menu: how big the guide's page is drawn on the screen (the guide's page only; the menus and the tools keep their size). The commands are dimmed with no guide open.
  const zoom = ctx.zoom;
  /** @type {MenuDef} */
  const view = {
    id: "view",
    label: "View",
    items: [
      { label: "Zoom in", icon: "zoom-in", keys: "Ctrl+Alt+=", title: "Draw the guide's page bigger on the screen (the mouse wheel with Ctrl held, over the guide, does the same)", run: () => void zoom?.more(), enabled: () => zoom?.can() === true },
      { label: "Zoom out", icon: "zoom-out", keys: "Ctrl+Alt+-", title: "Draw the guide's page smaller on the screen", run: () => void zoom?.less(), enabled: () => zoom?.can() === true },
      { label: "Actual size", icon: "rotate-ccw", keys: "Ctrl+Alt+0", title: "Draw the guide's page at its usual size, 100%", run: () => void zoom?.actual(), enabled: () => zoom?.can() === true },
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
      { label: "Check for updates...", icon: "refresh-cw", title: "Whether a newer OpenStage Guide is published. It asks GitHub only when you choose this, and installs nothing", run: ctx.updates ?? (() => undefined) },
      { label: "About OpenStage Guide", icon: "info", title: "What this program is, its version, its license and what it is made with", run: ctx.about },
    ],
  };
  const middle = PROGRAM_MENUS.map((m) => /** @type {MenuDef} */ ({ id: m.id, label: m.label, items: m.tools.map(tool) }));
  /** @type {MenuCommand[]} */
  const quick = [
    { label: "Review guide", icon: "list-checks", title: "Find missing text, objectives, unfinished prompts and timing issues", run: () => ctx.openTool("reviewGuide") },
    { label: "Update guide", icon: "refresh-cw", title: "Renumber the headings and forms, and correct the times", run: ctx.updateGuide },
    { label: "Make workbook", icon: "book-open", title: "Make the participant workbook", run: () => ctx.openTool("participantWorkbook") },
  ];
  return { menus: [file, edit, insert, format, view, ...middle, help], quick };
}

export function menuEntries(/** @type {MenuDef[]} */ menus, /** @type {string} */ id) {
  return menus.find((m) => m.id === id)?.items ?? [];
}
