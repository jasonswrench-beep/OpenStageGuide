/**
 * The small pieces the design desk's screens are drawn with: a labelled box, a choice, a tick, a button, a table. Each makes its element with text and attributes only, never
 * markup, gives every control a label a screen reader can say, and calls `onChange` with the new value as it changes. The words are the author's own, so they are only ever set as text.
 */
import { el } from "./menubar.js";

let counter = 0;
/** A new id for a control, so its label can name it. @param {string} [prefix] */
export const uid = (prefix = "dd") => `${prefix}-${++counter}`;

/**
 * A box with its label above it (several lines when `rows` is more than 1).
 * @param {Document} doc
 * @param {{ label: string, value?: string, rows?: number, hint?: string, onInput?: (v: string) => void, type?: string, max?: number, small?: boolean, id?: string, placeholder?: string }} o
 */
export function field(doc, o) {
  const id = o.id ?? uid("f");
  const hintId = `${id}-hint`;
  const attrs = { id, ...(o.hint ? { "aria-describedby": hintId } : {}), ...(o.placeholder ? { placeholder: o.placeholder } : {}) };
  const input =
    (o.rows ?? 1) > 1
      ? el(doc, "textarea", { ...attrs, rows: String(o.rows), spellcheck: "true" })
      : el(doc, "input", { ...attrs, type: o.type ?? "text", autocomplete: "off", ...(o.type === "number" ? { inputmode: "numeric", min: "0", max: String(o.max ?? 999) } : { maxlength: String(o.max ?? 300) }) });
  input.value = o.value ?? "";
  if (o.onInput) input.addEventListener("input", () => o.onInput?.(input.value));
  return Object.assign(el(doc, "div", { class: `field${o.small ? " small" : ""}` }, el(doc, "label", { class: "field-label", for: id, text: o.label }), input, ...(o.hint ? [el(doc, "p", { id: hintId, class: "hint", text: o.hint })] : [])), { input });
}

/**
 * A choice from a list, with its label above it.
 * @param {Document} doc
 * @param {{ label: string, value?: string, options: { value: string, label: string }[], onChange?: (v: string) => void, small?: boolean, hint?: string, id?: string }} o
 */
export function choice(doc, o) {
  const id = o.id ?? uid("s");
  const hintId = `${id}-hint`;
  const select = el(doc, "select", { id, ...(o.hint ? { "aria-describedby": hintId } : {}) }, ...o.options.map((x) => el(doc, "option", { value: x.value, text: x.label })));
  select.value = o.value ?? o.options[0]?.value ?? "";
  if (o.onChange) select.addEventListener("change", () => o.onChange?.(select.value));
  return Object.assign(el(doc, "div", { class: `field${o.small ? " small" : ""}` }, el(doc, "label", { class: "field-label", for: id, text: o.label }), select, ...(o.hint ? [el(doc, "p", { id: hintId, class: "hint", text: o.hint })] : [])), { select });
}

/**
 * A tick with its words beside it, and a quiet line under them when there is one.
 * @param {Document} doc
 * @param {{ label: string, checked?: boolean, note?: string, onChange?: (on: boolean) => void }} o
 */
export function tick(doc, o) {
  const id = uid("c");
  const box = el(doc, "input", { type: "checkbox", id, ...(o.note ? { "aria-describedby": `${id}-note` } : {}) });
  box.checked = o.checked === true;
  if (o.onChange) box.addEventListener("change", () => o.onChange?.(box.checked));
  return Object.assign(el(doc, "div", { class: "check" }, box, el(doc, "label", { for: id }, el(doc, "span", { text: o.label }), ...(o.note ? [el(doc, "span", { id: `${id}-note`, class: "hint", text: o.note })] : []))), { box });
}

/**
 * A button. `label` is what a screen reader says when the words on it are not enough (a button that only says "Remove" in a list of ten).
 * @param {Document} doc
 * @param {string} text @param {() => unknown} onClick
 * @param {{ primary?: boolean, danger?: boolean, small?: boolean, label?: string, title?: string, disabled?: boolean }} [o]
 */
export function button(doc, text, onClick, o = {}) {
  const b = el(doc, "button", {
    type: "button",
    text,
    ...(o.primary || o.danger || o.small ? { class: [o.primary ? "primary" : "", o.danger ? "danger" : "", o.small ? "small-button" : ""].filter(Boolean).join(" ") } : {}),
    ...(o.label ? { "aria-label": o.label } : {}),
    ...(o.title ? { title: o.title } : {}),
    ...(o.disabled ? { disabled: "" } : {}),
  });
  b.addEventListener("click", () => {
    void Promise.resolve(onClick()).catch(() => {});
  });
  return b;
}

/** A row of controls. @param {Document} doc @param {...(Node | string)} kids */
export const row = (doc, ...kids) => el(doc, "div", { class: "dd-row" }, ...kids);
/** A heading of the desk. @param {Document} doc @param {number} level @param {string} text */
export const heading = (doc, level, text) => el(doc, `h${level}`, { text });
/** A quiet line. @param {Document} doc @param {string} text */
export const note = (doc, text) => el(doc, "p", { class: "hint", text });

/**
 * A table of words, with its caption (a screen reader says what it is) and a header row.
 * @param {Document} doc
 * @param {{ caption: string, head: string[], body: string[][] }} t
 */
export function table(doc, t) {
  return el(
    doc,
    "table",
    { class: "dd-table" },
    el(doc, "caption", { text: t.caption }),
    el(doc, "thead", {}, el(doc, "tr", {}, ...t.head.map((h) => el(doc, "th", { scope: "col", text: h })))),
    el(doc, "tbody", {}, ...t.body.map((r) => el(doc, "tr", {}, ...r.map((c, i) => el(doc, i === 0 ? "th" : "td", { ...(i === 0 ? { scope: "row" } : {}), text: c }))))),
  );
}

/**
 * Several ticks that belong together, in a group a screen reader names.
 * @param {Document} doc
 * @param {{ legend: string, items: { id: string, label: string, note?: string }[], chosen: string[], onChange: (ids: string[]) => void, empty?: string }} o
 */
export function tickGroup(doc, o) {
  const set = new Set(o.chosen);
  const fs = el(doc, "fieldset", { class: "dd-group" }, el(doc, "legend", { text: o.legend }));
  if (!o.items.length) fs.append(el(doc, "p", { class: "hint", text: o.empty ?? "None yet." }));
  for (const item of o.items)
    fs.append(
      tick(doc, {
        label: item.label,
        note: item.note,
        checked: set.has(item.id),
        onChange: (on) => {
          if (on) set.add(item.id);
          else set.delete(item.id);
          o.onChange(o.items.map((x) => x.id).filter((id) => set.has(id)));
        },
      }),
    );
  return fs;
}

/** Give a file to the person: the browser's own download. @param {Document} doc @param {Uint8Array | string} data @param {string} name @param {string} type */
export function give(doc, data, name, type) {
  const url = URL.createObjectURL(new Blob([/** @type {BlobPart} */ (data)], { type }));
  const a = el(doc, "a", { href: url, download: name, hidden: "" });
  doc.body.append(a);
  a.click();
  a.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30000);
}
