// @ts-check
/**
 * The small dialogs the program opens: a name for a new guide, a choice among the guides there are, and About. They follow the suite's dialog:
 * a Close button in the corner that does what Escape does, a heading, and a foot of buttons whose main one is filled. Focus starts on the first
 * thing to fill in or choose, and goes back to where it was when the dialog closes. Text and attributes only, never markup.
 */
import { ICONS_LINE, LICENSE_LINE, MADE_WITH, NOTICE_LINE, aboutVersion } from "./credits.js";
import { guideName } from "./guide-name.js";
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
    let working = false;
    const done = async () => {
      // A second press while the first is still being dealt with (a slow disk, a double press of Enter) does nothing.
      if (working) return;
      working = true;
      problem.hidden = true;
      try {
        result = await spec.finish(dialog);
        dialog.close();
      } catch (e) {
        problem.textContent = /** @type {Error} */ (e).message;
        problem.hidden = false;
      } finally {
        working = false;
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
    const first = spec.first?.(dialog) ?? confirm;
    first.focus();
    // A first box that already holds words has them chosen, so what the person types replaces them (it joined the new words on to the old ones).
    const box = /** @type {HTMLInputElement | HTMLTextAreaElement} */ (first);
    if ((box.tagName === "INPUT" && box.type === "text") || box.tagName === "TEXTAREA") {
      if (box.value) box.select();
    }
  });
}

/**
 * Ask for a name. Resolves to the name, or null when the dialog was closed.
 * `accept`, when given, is tried with the name before the dialog closes; if it throws, its message is shown in the dialog, which stays open.
 * @param {Document} doc @param {{ title: string, label: string, confirm: string, hint?: string, value?: string, accept?: (name: string) => Promise<unknown> | unknown }} opts
 */
export function askName(doc, opts) {
  /** @type {HTMLInputElement} */
  const input = el(doc, "input", { type: "text", id: "dlg-name", autocomplete: "off", spellcheck: "false", maxlength: "80", value: opts.value ?? "" });
  return openDialog(doc, {
    title: opts.title,
    body: () => [el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: "dlg-name", text: opts.label }), input, ...(opts.hint ? [el(doc, "p", { class: "hint", text: opts.hint })] : []))],
    confirm: opts.confirm,
    first: () => input,
    finish: async () => {
      const name = input.value.trim();
      if (!name) throw new Error("Type a name first.");
      // What is done with the name is tried here, so a refusal (a name Windows does not allow, or one in use) is told in the dialog, which stays
      // open with what was typed, instead of closing and leaving a message behind it.
      if (opts.accept) await opts.accept(name);
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
      el(doc, "ul", { class: "dialog-list", role: "radiogroup", "aria-label": "Guides" }, ...guides.map((g, i) => el(doc, "li", {}, el(doc, "label", {}, el(doc, "input", { type: "radio", name: "dlg-guide", value: g.name, ...(i === 0 ? { checked: "" } : {}) }), el(doc, "span", { class: "recent-name", text: guideName(g.name) }), el(doc, "span", { class: "recent-when", text: when(g.when) }))))),
    ],
    confirm: "Open",
    first: (d) => /** @type {HTMLElement | null} */ (d.querySelector("input[type=radio]:checked")),
    finish: (d) => /** @type {HTMLInputElement | null} */ (d.querySelector("input[type=radio]:checked"))?.value ?? null,
  });
}

/**
 * Choose one of a few things, each with a name and a line about it. Resolves to the id of the one chosen, or null.
 * @param {Document} doc @param {{ title: string, intro?: string, confirm: string, options: { id: string, label: string, note?: string }[], current?: string }} opts
 */
export function chooseOne(doc, opts) {
  return openDialog(doc, {
    title: opts.title,
    intro: opts.intro,
    body: () => [
      el(
        doc,
        "ul",
        { class: "dialog-list", role: "radiogroup", "aria-label": opts.title },
        ...opts.options.map((o, i) =>
          el(
            doc,
            "li",
            {},
            el(
              doc,
              "label",
              {},
              el(doc, "input", { type: "radio", name: "dlg-choice", value: o.id, ...((opts.current ? o.id === opts.current : i === 0) ? { checked: "" } : {}) }),
              el(doc, "span", {}, el(doc, "strong", { text: o.label }), ...(o.note ? [el(doc, "span", { class: "hint", text: ` ${o.note}` })] : [])),
            ),
          ),
        ),
      ),
    ],
    confirm: opts.confirm,
    first: (d) => /** @type {HTMLElement | null} */ (d.querySelector("input[type=radio]:checked")),
    finish: (d) => /** @type {HTMLInputElement | null} */ (d.querySelector("input[type=radio]:checked"))?.value ?? null,
  });
}

/**
 * A dialog of fields to fill in. Each field has an id and a label, and a type: "text", "number", "lines" (a box for several lines), "select"
 * (with options), "check" or "file" (a file to choose; its value is the File, or null). `finish` gets the values by id and returns the result
 * (it may be a promise), or throws a message to show in the dialog and keep it open. Resolves to the result, or null when the dialog was closed.
 * @template T
 * @param {Document} doc
 * @param {{ title: string, intro?: string, confirm: string, fields: { id: string, label: string, type: "text" | "number" | "lines" | "select" | "check" | "file", value?: string | boolean, hint?: string, rows?: number, accept?: string, options?: { value: string, label: string }[], list?: { value: string, label: string }[] }[], finish: (values: Record<string, string | boolean | File | null>) => T | Promise<T> }} opts
 */
export function askFields(doc, opts) {
  /** @type {Map<string, HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>} */
  const inputs = new Map();
  /** @type {Node[]} */
  const rows = [];
  for (const f of opts.fields) {
    const id = `dlg-f-${f.id}`;
    /** @type {HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement} */
    let input;
    if (f.type === "lines") {
      input = el(doc, "textarea", { id, rows: String(f.rows ?? 4), spellcheck: "true" });
      input.value = String(f.value ?? "");
    } else if (f.type === "select") {
      input = el(doc, "select", { id }, ...(f.options ?? []).map((o) => el(doc, "option", { value: o.value, text: o.label })));
      input.value = String(f.value ?? f.options?.[0]?.value ?? "");
    } else if (f.type === "file") {
      input = el(doc, "input", { type: "file", id, ...(f.accept ? { accept: f.accept } : {}) });
    } else if (f.type === "check") {
      input = el(doc, "input", { type: "checkbox", id });
      /** @type {HTMLInputElement} */ (input).checked = f.value === true;
    } else {
      input = el(doc, "input", { type: "text", id, autocomplete: "off", ...(f.type === "number" ? { inputmode: "numeric", maxlength: "3" } : { maxlength: "200" }), ...(f.list ? { list: `${id}-list` } : {}), value: String(f.value ?? "") });
    }
    // Choices to start from (any text can still be typed).
    if (f.list) rows.push(el(doc, "datalist", { id: `${id}-list` }, ...f.list.map((o) => el(doc, "option", { value: o.value, label: o.label }))));
    inputs.set(f.id, input);
    if (f.hint) input.setAttribute("aria-describedby", `${id}-hint`);
    const hint = f.hint ? [el(doc, "p", { id: `${id}-hint`, class: "hint", text: f.hint })] : [];
    rows.push(
      f.type === "check"
        ? el(doc, "div", { class: "field" }, el(doc, "label", { class: "check-label", for: id }, input, el(doc, "span", { text: ` ${f.label}` })), ...hint)
        : el(doc, "div", { class: "field" }, el(doc, "label", { class: "field-label", for: id, text: f.label }), input, ...hint),
    );
  }
  return openDialog(doc, {
    title: opts.title,
    intro: opts.intro,
    body: () => rows,
    confirm: opts.confirm,
    first: () => inputs.values().next().value ?? null,
    finish: () => {
      /** @type {Record<string, string | boolean | File | null>} */
      const values = {};
      for (const [id, input] of inputs) {
        if (input instanceof HTMLInputElement && input.type === "checkbox") values[id] = input.checked;
        else if (input instanceof HTMLInputElement && input.type === "file") values[id] = input.files?.[0] ?? null;
        else values[id] = input.value;
      }
      return opts.finish(values);
    },
  });
}

/**
 * The About window: what the program is, which version it is (as the server reports it), its license and what it is made with.
 * @param {Document} doc @param {{ version?: unknown, build?: unknown } | null} [info]
 */
export function aboutDialog(doc, info) {
  const version = aboutVersion(info);
  return openDialog(doc, {
    title: "About OpenStage Guide",
    body: () => [
      ...(version ? [el(doc, "p", { class: "about-version", text: version })] : []),
      el(doc, "p", { text: "OpenStage Guide makes facilitator guides and participant workbooks: the script, the forms and handouts around it, the timing and the checks." }),
      el(doc, "p", { class: "hint", text: "This is the first form of the program: the guide is kept in a file on this computer, and the tools beside it are the ones the Word add-in has. Nothing leaves this computer." }),
      el(doc, "h3", { class: "about-head", text: "License" }),
      el(doc, "p", { text: LICENSE_LINE }),
      el(doc, "p", { text: ICONS_LINE }),
      el(doc, "h3", { class: "about-head", text: "Made with" }),
      el(doc, "ul", { class: "about-list" }, ...MADE_WITH.map((c) => el(doc, "li", { text: `${c.name}: ${c.use}. ${c.license}.` }))),
      el(doc, "p", { class: "hint", text: NOTICE_LINE }),
    ],
    confirm: "Close",
    cancel: "",
    finish: () => true,
  });
}
