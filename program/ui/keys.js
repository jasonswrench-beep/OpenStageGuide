// @ts-check
/**
 * Every key the program has, in groups, for the Keyboard keys window (Help > Keyboard keys, Ctrl+/), and the window itself. The look
 * follows OpenStage Author's window of the same name by way of OpenAssess's copy (see NOTICE): a table for each group, each named by its heading.
 */

import { el } from "./menubar.js";

/** @typedef {{ keys: string, does: string }} KeyRow */
/** @typedef {{ id: string, title: string, rows: KeyRow[] }} KeyGroup */

/** @type {KeyGroup[]} */
export const KEY_GROUPS = [
  {
    id: "anywhere",
    title: "Anywhere",
    rows: [
      { keys: "Ctrl+K", does: "Search for a command by its name" },
      { keys: "Ctrl+/", does: "This list of keys" },
      { keys: "F10, or Alt pressed and let go on its own", does: "Go to the menus, and pressed again, back to where you were" },
      { keys: "Tab, Shift+Tab", does: "The next or the last control; the first Tab stop on a screen is the link that skips to its content" },
    ],
  },
  {
    id: "menus",
    title: "Menus, and the ... buttons",
    rows: [
      { keys: "Down, Enter or Space", does: "Open the menu on its first item" },
      { keys: "Up", does: "Open the menu on its last item" },
      { keys: "Up, Down, Home, End", does: "Move in a menu" },
      { keys: "A letter", does: "The next item that starts with it" },
      { keys: "Left, Right", does: "In the menus under the navy bar: the menu beside it" },
      { keys: "Escape, Tab", does: "Close the menu and go back to where it was opened from" },
    ],
  },
  {
    id: "guide",
    title: "The guide and its tools",
    rows: [
      { keys: "Tab", does: "Into the guide, from it to the tools beside it, and on" },
      { keys: "Up, Down, Home, End", does: "Move between the paragraphs of the guide" },
      { keys: "Enter, F2, or a click", does: "Change the words of a paragraph (the time line, tables, pictures and check boxes are changed with the tools)" },
      { keys: "Escape", does: "Stop changing the words; with paragraphs chosen, stop choosing" },
      { keys: "Shift+Down, Shift+Up", does: "Choose paragraphs together, a row at a time" },
      { keys: "Ctrl+A", does: "Choose every paragraph" },
      { keys: "Alt+Up, Alt+Down", does: "Move the paragraph you are on, or the ones chosen, up or down" },
      { keys: "Delete", does: "Delete the paragraph you are on, or the ones chosen (Ctrl+Z brings them back)" },
      { keys: "Ctrl+Enter", does: "A new paragraph below, ready to type in" },
      { keys: "Ctrl+C, Ctrl+X, Ctrl+V", does: "Copy, cut and paste paragraphs; several lines pasted into the words become a paragraph each" },
      { keys: "Page Up, Page Down", does: "Scroll the guide" },
    ],
  },
  {
    id: "words",
    title: "Changing the words of a paragraph",
    rows: [
      { keys: "Enter", does: "Start the next paragraph with the words after the caret (in a heading: go to the next row)" },
      { keys: "Backspace at the start", does: "Join the paragraph to the one before; an empty paragraph is taken away" },
      { keys: "Up or Left at the start, Down or Right at the end", does: "Go to the paragraph before or after" },
      { keys: "Ctrl+Z, Ctrl+Y (or Ctrl+Shift+Z)", does: "Undo and Redo: one history for typing and for what the tools add" },
    ],
  },
  {
    id: "search",
    title: "Search for a command (Ctrl+K)",
    rows: [
      { keys: "Type", does: "Find the commands whose names have those words" },
      { keys: "Up, Down, Page Up, Page Down", does: "Move through what is found" },
      { keys: "Enter", does: "Do the command" },
      { keys: "Escape", does: "Close the search" },
    ],
  },
  {
    id: "dialogs",
    title: "Dialogs",
    rows: [
      { keys: "Tab, Shift+Tab", does: "Move through the dialog" },
      { keys: "Enter", does: "Do the dialog's main button, when a box for typing has the keyboard" },
      { keys: "Escape", does: "Close the dialog without changing anything" },
    ],
  },
];

/** Only one at a time. @type {HTMLDialogElement | null} */
let openNow = null;

/**
 * The Keyboard keys window. Focus goes back to where it was when it closes.
 * @param {Document} doc
 */
export function openKeys(doc) {
  if (openNow) {
    openNow.querySelector("button")?.focus();
    return;
  }
  const from = /** @type {HTMLElement | null} */ (doc.activeElement);
  /** @type {HTMLDialogElement} */
  const dialog = el(doc, "dialog", { class: "keys-dialog", "aria-labelledby": "ky-title", "aria-describedby": "ky-intro" });
  const close = el(doc, "button", { type: "button", class: "primary", text: "Close" });
  close.addEventListener("click", () => dialog.close());
  dialog.append(
    el(
      doc,
      "div",
      { class: "dialog-body" },
      el(doc, "h2", { id: "ky-title", text: "Keyboard keys" }),
      el(doc, "p", { id: "ky-intro", class: "hint", text: "Every key OpenStage Guide has. The menus show each command's keys beside it, and Ctrl+K finds a command by its name." }),
      ...KEY_GROUPS.map((g) =>
        el(
          doc,
          "section",
          { class: "keys-group", "aria-labelledby": `ky-${g.id}` },
          el(doc, "h3", { id: `ky-${g.id}`, text: g.title }),
          el(
            doc,
            "table",
            { class: "keys-table", "aria-labelledby": `ky-${g.id}` },
            el(doc, "thead", {}, el(doc, "tr", {}, el(doc, "th", { scope: "col", text: "Keys" }), el(doc, "th", { scope: "col", text: "What they do" }))),
            el(doc, "tbody", {}, ...g.rows.map((r) => el(doc, "tr", {}, el(doc, "th", { scope: "row", text: r.keys }), el(doc, "td", { text: r.does })))),
          ),
        ),
      ),
      el(doc, "div", { class: "dialog-actions" }, close),
    ),
  );
  dialog.addEventListener("close", () => {
    dialog.remove();
    openNow = null;
    if (from && from !== doc.body && from.isConnected && !from.closest("[hidden]")) from.focus();
  });
  doc.body.append(dialog);
  openNow = dialog;
  dialog.showModal();
  close.focus();
}
