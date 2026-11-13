// @ts-check
/**
 * The box a paragraph's words are changed in, which shows each stretch with its own look (bold, italic, color...) while it is typed in. It
 * stands where the plain text box stood and answers the same way (value, selectionStart, selectionEnd, setSelectionRange, setRangeText, focus,
 * and the events input, keydown, paste and focusout), so the editor's keys and saving did not change.
 *
 * What the box holds is a list of spans (src/guide-runs.mjs): the words and the Word run properties of each stretch. The page's elements are
 * only a drawing of that list. The browser types, deletes and moves the caret in them as it does in any box; after each change the list is read
 * back from what the browser left (only the words, and the look each stretch was drawn with), and the elements are drawn again when the browser
 * has left them in a shape the box would not have made. Everything that could bring markup in from outside (the browser's own bold keys, paste,
 * dropping text, the browser's own undo, a new line) is refused, so the only way words get a look is a command of the program. Text and
 * attributes only, never markup.
 */
import { applyChange, fmtOfRpr, normalize, rprAt, replaceText, runCss, summarize, textOf, withFmt } from "../../src/guide-runs.mjs";
import { drawCss } from "./preview.js";

const NBSP = /\u00a0/g;
const LINE_ENDS = /\r\n|[\r\n\u2028\u2029]/g;

/**
 * @typedef {{ text: string, rpr: string }} Span
 * @typedef {{ parse: (xml: string) => Document, serialize: (el: Element) => string }} Io
 */

/**
 * @param {Document} doc
 * @param {{ spans: Span[], label: string, describedby?: string, io: Io, styleCss?: (id: string) => Record<string, string> | null }} opts
 */
export function createRichBox(doc, { spans, label, describedby, io, styleCss }) {
  const win = /** @type {Window} */ (doc.defaultView);
  const el = doc.createElement("div");
  el.className = "g-edit g-rich";
  el.contentEditable = "true";
  el.spellcheck = true;
  el.setAttribute("role", "textbox");
  el.setAttribute("aria-multiline", "false");
  el.setAttribute("aria-label", label);
  if (describedby) el.setAttribute("aria-describedby", describedby);

  /** @type {Span[]} */
  let model = normalize(spans);
  /** The properties text of each look the box has drawn, by the number its elements carry. @type {string[]} */
  const table = [];
  /** A look chosen with no words chosen, kept for the next words typed there. @type {{ at: number, change: Record<string, unknown> } | null} */
  let pending = null;
  let composing = false;
  /** Where the caret was last, for when the box does not have the focus. @type {[number, number]} */
  let lastSel = [0, 0];

  const keyOf = (/** @type {string} */ rpr) => {
    let i = table.indexOf(rpr);
    if (i < 0) i = table.push(rpr) - 1;
    return i;
  };

  /** @param {Span} s */
  function spanElement(s) {
    const e = doc.createElement("span");
    e.dataset.k = String(keyOf(s.rpr));
    e.textContent = s.text;
    // A character style's look first, then what was given to these words themselves.
    const fmt = fmtOfRpr(s.rpr, io);
    const css = { ...(fmt.style ? styleCss?.(fmt.style) ?? {} : {}), ...runCss(fmt) };
    if (Object.keys(css).length) drawCss(e, css);
    return e;
  }

  /** Draw the list again, and put the caret back. @param {number} [a] @param {number} [b] */
  function render(a, b) {
    el.replaceChildren(...model.map(spanElement));
    if (a !== undefined) setSelection(a, b ?? a);
  }

  /** The text offset of a place in the box. @param {Node} node @param {number} offset */
  function offsetOf(node, offset) {
    const r = doc.createRange();
    r.setStart(el, 0);
    r.setEnd(node, offset);
    return r.toString().length;
  }

  /** The place in the box for a text offset: the end of the words before it where two stretches meet, so typing goes on in the one before. @param {number} offset @returns {[Node, number]} */
  function position(offset) {
    const walker = doc.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let left = Math.max(0, offset);
    /** @type {Text | null} */
    let last = null;
    while (walker.nextNode()) {
      const t = /** @type {Text} */ (walker.currentNode);
      if (left <= t.data.length) return [t, left];
      left -= t.data.length;
      last = t;
    }
    return last ? [last, last.data.length] : [el, 0];
  }

  /** @param {number} a @param {number} [b] */
  function setSelection(a, b = a) {
    const sel = win.getSelection();
    if (!sel) return;
    const [n1, o1] = position(a);
    const [n2, o2] = position(b);
    const r = doc.createRange();
    r.setStart(n1, o1);
    r.setEnd(n2, o2);
    sel.removeAllRanges();
    sel.addRange(r);
    lastSel = [a, b];
  }

  /** @returns {[number, number]} */
  function offsets() {
    const sel = win.getSelection();
    if (!sel || sel.rangeCount === 0 || !sel.anchorNode || !el.contains(sel.anchorNode) || !sel.focusNode || !el.contains(sel.focusNode)) return lastSel;
    const r = sel.getRangeAt(0);
    lastSel = [offsetOf(r.startContainer, r.startOffset), offsetOf(r.endContainer, r.endOffset)];
    return lastSel;
  }

  /** The list, read back from what the browser left in the box. @returns {Span[]} */
  function readModel() {
    /** @type {Span[]} */
    const out = [];
    let current = model[0]?.rpr ?? "";
    const walk = (/** @type {Node} */ node, /** @type {string | null} */ inherited) => {
      for (const child of Array.from(node.childNodes)) {
        if (child.nodeType === 3) {
          out.push({ text: /** @type {Text} */ (child).data.replace(NBSP, " ").replace(LINE_ENDS, " "), rpr: inherited ?? current });
        } else if (child.nodeType === 1) {
          const c = /** @type {HTMLElement} */ (child);
          if (c.tagName === "BR") continue;
          const k = c.dataset?.k;
          const rpr = k !== undefined && table[Number(k)] !== undefined ? table[Number(k)] : inherited;
          walk(c, rpr);
          if (rpr !== null) current = rpr;
        }
      }
    };
    walk(el, null);
    return normalize(out);
  }

  /** Whether the elements are what the box would have drawn from its list: spans, each with one text. */
  function canonical() {
    const kids = Array.from(el.childNodes);
    return kids.length === model.length && kids.every((n, i) => n.nodeType === 1 && /** @type {Element} */ (n).tagName === "SPAN" && n.childNodes.length === 1 && n.firstChild?.nodeType === 3 && /** @type {Text} */ (n.firstChild).data === model[i].text);
  }

  function announce() {
    el.dispatchEvent(new Event("input", { bubbles: true }));
  }

  /**
   * Other words in a stretch, the new ones with the look of the character before them (or `rpr`), and the caret after them.
   * @param {number} start @param {number} end @param {string} text @param {string} [rpr]
   */
  function edit(start, end, text, rpr) {
    const clean = text.replace(LINE_ENDS, " ");
    model = replaceText(model, start, end, clean, rpr);
    render(start + clean.length);
    announce();
  }

  // What the browser did to the words (typing, deleting, a spelling correction): read it back.
  el.addEventListener("input", () => {
    const read = readModel();
    model = read;
    if (!composing && !canonical()) {
      const [a, b] = offsets();
      render(a, b);
    }
  });
  el.addEventListener("compositionstart", () => {
    composing = true;
  });
  el.addEventListener("compositionend", () => {
    composing = false;
    model = readModel();
    if (!canonical()) {
      const [a, b] = offsets();
      render(a, b);
    }
  });

  // What the box refuses: anything that would give words a look, or a second line, other than the program's own commands.
  el.addEventListener("beforeinput", (ev) => {
    const type = /** @type {InputEvent} */ (ev).inputType;
    if (type.startsWith("format") || type === "historyUndo" || type === "historyRedo" || type === "insertParagraph" || type === "insertLineBreak" || type === "insertFromDrop" || type === "insertFromYank" || type === "insertFromPasteAsQuotation") {
      ev.preventDefault();
      return;
    }
    const data = /** @type {InputEvent} */ (ev).data;
    if (type === "insertText" && data && pending && !/** @type {InputEvent} */ (ev).isComposing) {
      const [a, b] = offsets();
      if (pending.at === a) {
        ev.preventDefault();
        const look = withFmt(rprAt(model, a), /** @type {any} */ (pending.change), io);
        pending = null;
        edit(a, b, data, look);
        return;
      }
      pending = null;
    }
  });
  // Paste is the box's own: the words as plain text. (Several lines are left to the editor, which makes paragraphs of them.)
  el.addEventListener("paste", (ev) => {
    ev.preventDefault();
    const raw = ev.clipboardData?.getData("text/plain") ?? "";
    if (raw.split(/\r\n|[\r\n]/).filter((l) => l.trim()).length >= 2) return;
    const [a, b] = offsets();
    edit(a, b, raw);
  });
  el.addEventListener("drop", (ev) => ev.preventDefault());

  const box = {
    el,
    /** The words, without their looks. */
    get value() {
      return textOf(model);
    },
    set value(/** @type {string} */ v) {
      model = normalize([{ text: v.replace(LINE_ENDS, " "), rpr: rprAt(model, 0) }]);
      render(v.length);
    },
    get selectionStart() {
      return offsets()[0];
    },
    get selectionEnd() {
      return offsets()[1];
    },
    /** @param {number} a @param {number} [b] */
    setSelectionRange(a, b = a) {
      setSelection(a, b);
    },
    /**
     * @param {string} text @param {number} [start] @param {number} [end] @param {"end" | "preserve"} [mode]
     */
    setRangeText(text, start, end, mode) {
      const [a, b] = offsets();
      const s = start ?? a;
      const e = end ?? b;
      edit(s, e, text);
      if (mode === "preserve") setSelection(s, s);
    },
    /** The stretches with their looks. @returns {Span[]} */
    spans() {
      return model.map((s) => ({ ...s }));
    },
    focus() {
      el.focus({ preventScroll: true });
    },
    /**
     * Give the words that are chosen a look (a change from guide-runs.mjs). With none chosen, the next words typed here have it.
     * @param {Record<string, unknown>} change @returns {boolean} whether words were changed
     */
    apply(change) {
      const [a, b] = offsets();
      if (a === b) {
        pending = { at: a, change: { ...(pending && pending.at === a ? pending.change : {}), ...change } };
        return false;
      }
      model = applyChange(model, a, b, /** @type {any} */ (change), io);
      render(a, b);
      announce();
      return true;
    },
    /** What the chosen words are like, setting by setting (the value they all have, 'mixed', or undefined), with a look waiting for the next words counted. */
    state() {
      const [a, b] = offsets();
      const s = /** @type {Record<string, unknown>} */ (summarize(model, a, b, io));
      return pending && pending.at === a && a === b ? { ...s, ...pending.change } : s;
    },
    /** Start from other spans, with the caret where it was or at `at`. @param {Span[]} spans @param {number} [at] */
    setSpans(spans, at) {
      model = normalize(spans);
      render(at ?? offsets()[0]);
    },
  };
  // The plain text box's answers, on the element too (what is asked of whatever has the focus).
  Object.defineProperties(el, {
    value: { get: () => box.value, configurable: true },
    selectionStart: { get: () => box.selectionStart, configurable: true },
    selectionEnd: { get: () => box.selectionEnd, configurable: true },
  });
  render();
  return box;
}
