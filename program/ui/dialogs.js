// @ts-check
/**
 * The small dialogs the program opens: a name for a new guide, a choice among the guides there are, and About. They follow the suite's dialog:
 * a Close button in the corner that does what Escape does, a heading, and a foot of buttons whose main one is filled. Focus starts on the first
 * thing to fill in or choose, and goes back to where it was when the dialog closes. Text and attributes only, never markup.
 */
import { el, uiIcon } from "./menubar.js";

/**
 * Open a dialog and wait for it to close. `body` makes its contents; `finish` is asked when the main button is chosen and returns the result
 * (or throws a message to show in the dialog and keep it open).
 * @template T
 * @param {Document} doc
 * @param {{ title: string, intro?: string, body?: (dialog: HTMLDialogElement) => Node[], confirm: string, cancel?: string, first?: (dialog: HTMLDialogElement) => HTMLElement | null, finish: (dialog: HTMLDialogElement) => T | Promise<T> }} spec
 * @returns {Promise<T | null>}
 */
export function openDialog(doc, spec) {
  return new Promise((resolve) => {
    const from = /** @type {HTMLElement | null} */ (doc.activeElement);
    const id = `dlg-${Math.random().toString(36).slice(2, 8)}`;
    /** @type {HTMLDialogElement} */
    const dialog = el(doc, "dialog", { "aria-labelledby": `${id}-title`, ...(spec.intro ? { "aria-describedby": `${id}-intro` } : {}) });
    const closeX = el(doc, "button", { type: "button", class: "dlg-close", "aria-label": "Close", title: "Close (Escape)" }, uiIcon(doc, "x"));
    const problem = el(doc, "p", { class: "problem", role: "alert", hidden: "" });
    const confirm = el(doc, "button", { type: "button", class: "primary", text: spec.confirm });
    const cancel = spec.cancel === "" ? null : el(doc, "button", { type: "button", text: spec.cancel ?? "Cancel" });
    let result = /** @type {T | null} */ (null);
    const done = async () => {
      problem.hidden = true;
      try {
        result = await spec.finish(dialog);
        dialog.close();
      } catch (e) {
        problem.textContent = /** @type {Error} */ (e).message;
        problem.hidden = false;
      }
    };
    confirm.addEventListener("click", done);
    cancel?.addEventListener("click", () => dialog.close());
    closeX.addEventListener("click", () => dialog.close());
    dialog.append(
      closeX,
      el(doc, "div", { class: "dialog-body" }, el(doc, "h2", { id: `${id}-title`, text: spec.title }), ...(spec.intro ? [el(doc, "p", { id: `${id}-intro`, class: "hint", text: spec.intro })] : []), ...(spec.body ? spec.body(dialog) : []), problem, el(doc, "div", { class: "dialog-actions" }, ...(cancel ? [cancel] : []), confirm)),
    );
    dialog.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && /** @type {HTMLElement} */ (e.target).tagName === "INPUT" && /** @type {HTMLInputElement} */ (e.target).type === "text") {
        e.preventDefault();
        done();
      }
    });
    dialog.addEventListener("close", () => {
      dialog.remove();
      if (from && from !== doc.body && from.isConnected && !from.closest("[hidden]")) from.focus();
      resolve(result);
    });
    doc.body.append(dialog);
    dialog.showModal();
    (spec.first?.(dialog) ?? confirm).focus();
  });
}

/**
 * Ask for a name. Resolves to the name, or null when the dialog was closed.
 * @param {Document} doc @param {{ title: string, label: string, confirm: string, hint?: string, value?: string }} opts
 */
export function askName(doc, opts) {
  /** @type {HTMLInputElement} */
  const input = el(doc, "input", { type: "text", id: "dlg-name", autocomplete: "off", spellcheck: "false", maxlength: "80", value: opts.value ?? "" });
  return openDialog(doc, {
    title: opts.title,
    body: () => [el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: "dlg-name", text: opts.label }), input, ...(opts.hint ? [el(doc, "p", { class: "hint", text: opts.hint })] : []))],
    confirm: opts.confirm,
    first: () => input,
    finish: () => {
      const name = input.value.trim();
      if (!name) throw new Error("Type a name first.");
      return name;
    },
  });
}

/**
 * Choose among the guides there are. Resolves to the file's name as the server listed it, or null.
 * @param {Document} doc @param {{ name: string, when: string }[]} guides @param {(iso: string) => string} when
 */
export function chooseGuide(doc, guides, when) {
  if (!guides.length) return openDialog(doc, { title: "Open guide", body: () => [el(doc, "p", { text: "There are no guides yet. Choose New guide to make one." })], confirm: "OK", cancel: "Close", finish: () => null });
  return openDialog(doc, {
    title: "Open guide",
    intro: "Choose the guide to open.",
    body: () => [
      el(doc, "ul", { class: "dialog-list", role: "radiogroup", "aria-label": "Guides" }, ...guides.map((g, i) => el(doc, "li", {}, el(doc, "label", {}, el(doc, "input", { type: "radio", name: "dlg-guide", value: g.name, ...(i === 0 ? { checked: "" } : {}) }), el(doc, "span", { class: "recent-name", text: g.name }), el(doc, "span", { class: "recent-when", text: when(g.when) }))))),
    ],
    confirm: "Open",
    first: (d) => /** @type {HTMLElement | null} */ (d.querySelector("input[type=radio]:checked")),
    finish: (d) => /** @type {HTMLInputElement | null} */ (d.querySelector("input[type=radio]:checked"))?.value ?? null,
  });
}

/** @param {Document} doc */
export function aboutDialog(doc) {
  return openDialog(doc, {
    title: "About OpenStage Guide",
    body: () => [
      el(doc, "p", { text: "OpenStage Guide makes facilitator guides and participant workbooks: the script, the forms and handouts around it, the timing and the checks." }),
      el(doc, "p", { class: "hint", text: "This is the first form of the program: the guide is kept in a file on this computer, and the tools beside it are the ones the Word add-in has. Nothing leaves this computer." }),
    ],
    confirm: "Close",
    cancel: "",
    finish: () => true,
  });
}
