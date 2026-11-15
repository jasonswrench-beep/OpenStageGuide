// @ts-check
/**
 * The welcome page: what the program shows with no guide open. Its two main actions as large buttons (New guide... filled in the program's
 * colour, and Open guide...), and a card of the guides there are, each one a button. After the suite's other programs' welcome pages.
 */
import { guideName } from "./guide-name.js";
import { el, uiIcon } from "./menubar.js";

/**
 * @param {Document} doc
 * @param {HTMLElement} main
 * `hidden` is how many guides the list leaves off because it was cleared: they are still in the folder, and the Open guide window lists them.
 * @param {{ guides: { name: string, when: string }[], hidden?: number, folder: string, when: (iso: string) => string, onNew: () => void, onOpen: () => void, onOpenNamed: (name: string) => void, onClear?: () => void, onShowAll?: () => void }} opts
 */
export function drawWelcome(doc, main, opts) {
  const newButton = el(doc, "button", { type: "button", class: "primary", id: "welcome-new" }, uiIcon(doc, "file-plus"), el(doc, "span", { text: "New guide..." }));
  const openButton = el(doc, "button", { type: "button", id: "welcome-open" }, uiIcon(doc, "folder-open"), el(doc, "span", { text: "Open guide..." }));
  newButton.addEventListener("click", opts.onNew);
  openButton.addEventListener("click", opts.onOpen);
  const list = opts.guides.length
    ? el(
        doc,
        "ul",
        {},
        ...opts.guides.map((g) => {
          const button = el(doc, "button", { type: "button", "data-guide": g.name }, uiIcon(doc, "file-text"), el(doc, "span", { class: "recent-name", text: guideName(g.name) }), el(doc, "span", { class: "recent-when", text: opts.when(g.when) }));
          button.addEventListener("click", () => opts.onOpenNamed(g.name));
          return el(doc, "li", {}, button);
        }),
      )
    : el(doc, "p", { class: "recent-empty", text: opts.hidden ? "The list is cleared. Your guides are still in the folder: Open guide shows every one of them." : "You have no guides yet. Choose New guide to start one." });
  // Clearing only leaves guides off this list: nothing is deleted or moved. Another button brings them back.
  const clearButton = opts.guides.length && opts.onClear ? el(doc, "button", { type: "button", id: "welcome-clear", title: "Leave these guides off this list. Nothing is deleted: Open guide still lists them all." }, uiIcon(doc, "eraser"), el(doc, "span", { text: "Clear previous guides" })) : null;
  clearButton?.addEventListener("click", () => opts.onClear?.());
  const showButton = opts.hidden && opts.onShowAll ? el(doc, "button", { type: "button", id: "welcome-show-all" }, uiIcon(doc, "rotate-ccw"), el(doc, "span", { text: opts.hidden === 1 ? "Show the guide left off" : `Show the ${opts.hidden} guides left off` })) : null;
  showButton?.addEventListener("click", () => opts.onShowAll?.());
  main.replaceChildren(
    el(
      doc,
      "div",
      { class: "welcome" },
      el(doc, "h1", { id: "screen-title", tabindex: "-1", text: "OpenStage Guide" }),
      el(doc, "p", { class: "lead", text: "Make a facilitator guide and the participant workbook that goes with it: the script, the forms and handouts around it, the timing and the checks." }),
      el(doc, "div", { class: "welcome-actions" }, newButton, openButton),
      el(doc, "section", { class: "recent", "aria-labelledby": "recent-title" }, el(doc, "div", { class: "recent-head" }, el(doc, "h2", { id: "recent-title", tabindex: "-1", text: "Your guides" }), ...(clearButton ? [clearButton] : []), ...(showButton ? [showButton] : [])), list),
      ...(opts.folder ? [el(doc, "p", { class: "where", text: `Your guides are kept in a folder named "${opts.folder}".` })] : []),
    ),
  );
}
