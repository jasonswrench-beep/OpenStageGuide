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
 * @param {{ guides: { name: string, when: string }[], folder: string, when: (iso: string) => string, onNew: () => void, onOpen: () => void, onOpenNamed: (name: string) => void }} opts
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
    : el(doc, "p", { class: "recent-empty", text: "You have no guides yet. Choose New guide to start one." });
  main.replaceChildren(
    el(
      doc,
      "div",
      { class: "welcome" },
      el(doc, "h1", { id: "screen-title", tabindex: "-1", text: "OpenStage Guide" }),
      el(doc, "p", { class: "lead", text: "Make a facilitator guide and the participant workbook that goes with it: the script, the forms and handouts around it, the timing and the checks." }),
      el(doc, "div", { class: "welcome-actions" }, newButton, openButton),
      el(doc, "section", { class: "recent", "aria-labelledby": "recent-title" }, el(doc, "div", { class: "recent-head" }, el(doc, "h2", { id: "recent-title", text: "Your guides" })), list),
      ...(opts.folder ? [el(doc, "p", { class: "where", text: `Your guides are kept in ${opts.folder}.` })] : []),
    ),
  );
}
