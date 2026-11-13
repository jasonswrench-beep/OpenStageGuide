// @ts-check
/**
 * The guide as the author writes it. The guide is drawn as rows, one for each paragraph (a table is one row). The arrow keys move between
 * rows; Enter, F2 or a click starts changing a row's words, which become a text box that grows with them; Escape stops. Enter in the words
 * starts the next paragraph; Backspace at the start joins a paragraph to the one before; Up and Down (and Left and Right) at the ends of the
 * words go to the next or the last paragraph. Every change is made on the guide's document (guide-word.mjs, with its saving) through
 * src/guide-edit.mjs, which changes the words and nothing else in a paragraph, so what the editor cannot show as words (a picture, a check
 * box, a citation, the time line, a table) is never rewritten: those rows are read, and changed with the tools beside the guide.
 *
 * The design is in docs/EDITING-SURFACE-DESIGN-2026-11-13.md. The page writes text and attributes only, never markup.
 */
import { isHeading, joinWords, readParagraph, setWords, splitWords } from "../../src/guide-edit.mjs";
import { reconcile, rowKey, rowSignature, sameList, wordsKey, wordsLabel } from "./editor-model.js";
import { describeGuide, drawNodes } from "./preview.js";

/** @param {string} xml */
const parseXml = (xml) => new DOMParser().parseFromString(xml, "application/xml");
const io = { parse: parseXml, serialize: (/** @type {Element} */ el) => new XMLSerializer().serializeToString(el) };
/** How long after the author stops moving or typing the tools are told where the cursor is. */
const SELECTION_DELAY = 250;

/**
 * @typedef {{ xml?: string, text?: string, style?: string }} Entry
 * @typedef {{ sig: string, el: HTMLElement }} Row
 * @typedef {{ row: Row, ta: HTMLTextAreaElement, group: string }} Editing
 */

/**
 * @param {Document} doc
 * @param {HTMLElement} host where the guide is drawn
 * @param {any} guide the guide's document, with its saving
 * @param {{ say: (message: string, bad?: boolean) => void }} ui
 */
export function mountEditor(doc, host, guide, ui) {
  const page = doc.createElement("article");
  page.className = "page";
  page.setAttribute("aria-label", "The guide");
  const hint = doc.createElement("p");
  hint.id = "edit-hint";
  hint.className = "sr-only";
  hint.textContent = "Enter starts a new paragraph. Escape stops changing the words. Up at the start or Down at the end goes to the next paragraph. Control Z undoes.";
  const empty = doc.createElement("p");
  empty.className = "g-empty";
  empty.textContent = "This guide is empty. Choose Guide overview in the tools beside it to start it.";
  host.replaceChildren(page, hint);

  /** @type {Entry[]} */
  let entries = [];
  /** @type {string[]} */
  let entrySigs = [];
  /** @type {Row[]} */
  let shown = [];
  /** @type {Editing | null} */
  let editing = null;
  let groupCounter = 0;
  let queue = Promise.resolve();
  /** @type {number | null} */
  let selectionTimer = null;
  let gone = false;

  /** @param {Row} row */
  const idx = (row) => Number(row.el.dataset.index);
  /** @param {number} index */
  const rowAt = (index) => shown.find((r) => idx(r) === index) ?? null;
  /** The row for a paragraph, or the nearest one above it that is drawn (some paragraphs, such as the contents lines, have no row). @param {number} index */
  const rowNear = (index) => rowAt(index) ?? [...shown].reverse().find((r) => idx(r) <= index) ?? shown[0] ?? null;
  /** @param {Row} row */
  const entryOf = (row) => entries[idx(row)] ?? { xml: "", text: "", style: "Normal" };
  /** @param {unknown} e */
  const messageOf = (e) => (e instanceof Error ? e.message : String(e));

  function tellTools() {
    if (selectionTimer !== null) window.clearTimeout(selectionTimer);
    selectionTimer = window.setTimeout(() => {
      selectionTimer = null;
      try {
        guide.fireSelection();
      } catch {
        /* the tools may not be there yet */
      }
    }, SELECTION_DELAY);
  }

  /** @param {Row} row */
  function placeCursor(row) {
    try {
      guide.place(idx(row));
    } catch {
      /* the document may be closing */
    }
    tellTools();
  }

  /** @param {Row | null} row the row that gets the Tab stop */
  function setTabStop(row) {
    for (const r of shown) r.el.tabIndex = r === row ? 0 : -1;
  }

  /** @param {Row} row @param {Entry} entry */
  function decorate(row, entry) {
    const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
    row.el.dataset.editable = info.editable ? "true" : "false";
    row.el.setAttribute("aria-description", info.editable ? "Press Enter to change the words." : info.reason ?? "This part is changed with the tools beside the guide.");
  }

  /** Read the document and bring the rows up to date: only what changed is drawn again. */
  function render() {
    if (gone) return;
    entries = guide.exportState().paragraphs;
    const sigs = entries.map(rowSignature);
    if (sameList(sigs, entrySigs)) return;
    entrySigs = sigs;
    const nodes = describeGuide(entries, parseXml, { blank: true });
    const rowSigs = nodes.map((n) => rowSignature(entries[n.index ?? 0]));
    const { prefix, oldEnd, newEnd } = reconcile(shown.map((r) => r.sig), rowSigs);
    // A row that is being replaced cannot go on being edited: the document changed under it (Undo, or the tools).
    if (editing && shown.indexOf(editing.row) >= prefix && shown.indexOf(editing.row) < oldEnd) editing = null;
    const hadFocus = doc.activeElement && page.contains(doc.activeElement);
    for (const r of shown.slice(prefix, oldEnd)) r.el.remove();
    if (empty.parentNode) empty.remove();
    const before = shown[oldEnd]?.el ?? null;
    /** @type {Row[]} */
    const made = [];
    for (let i = prefix; i < newEnd; i++) {
      const node = nodes[i];
      const tmp = doc.createElement("div");
      drawNodes(doc, tmp, [node]);
      const el = /** @type {HTMLElement} */ (tmp.firstElementChild);
      const row = { sig: rowSigs[i], el };
      el.tabIndex = -1;
      el.dataset.index = String(node.index);
      decorate(row, entries[node.index ?? 0]);
      page.insertBefore(el, before);
      made.push(row);
    }
    shown = [...shown.slice(0, prefix), ...made, ...shown.slice(oldEnd)];
    shown.forEach((r, i) => {
      r.el.dataset.index = String(nodes[i].index);
    });
    if (!shown.length) page.append(empty);
    if (!shown.some((r) => r.el.tabIndex === 0)) setTabStop(shown[0] ?? null);
    // The keyboard stays where it was, even when the row it was on has been drawn again.
    if (hadFocus && !doc.activeElement?.closest("#work-guide") && !editing) focusRow(rowNear(guide.cursor()));
  }

  /** @param {Row | null} row */
  function focusRow(row) {
    if (!row) return;
    setTabStop(row);
    row.el.focus();
    placeCursor(row);
  }

  /** Draw a row's own words again, after they were being changed. @param {Row} row */
  function redrawRow(row) {
    const entry = entryOf(row);
    const [node] = describeGuide([entry], parseXml, { blank: true });
    row.el.classList.remove("g-editing");
    if (!node) {
      row.el.replaceChildren();
      return;
    }
    const tmp = doc.createElement("div");
    drawNodes(doc, tmp, [node]);
    const fresh = tmp.firstElementChild;
    row.el.replaceChildren(...(fresh ? Array.from(fresh.childNodes) : []));
  }

  /** Stop changing words (they are already in the document), and draw the row as it reads. */
  function stopEdit() {
    if (!editing) return;
    const e = editing;
    editing = null;
    redrawRow(e.row);
  }

  /**
   * Change the words of a row.
   * @param {Row} row @param {number | "start" | "end"} caret
   */
  function startEdit(row, caret) {
    const entry = entryOf(row);
    const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!info.editable) {
      stopEdit();
      focusRow(row);
      ui.say(info.reason ?? "This part is changed with the tools beside the guide.");
      return;
    }
    if (editing?.row === row) {
      editing.ta.focus();
      return;
    }
    stopEdit();
    const ta = doc.createElement("textarea");
    ta.className = "g-edit";
    ta.rows = 1;
    ta.value = info.text;
    ta.spellcheck = true;
    ta.setAttribute("aria-label", wordsLabel(entry.style ?? "Normal"));
    ta.setAttribute("aria-describedby", "edit-hint");
    /** @type {HTMLElement[]} */
    const lead = [];
    if (info.prefix) {
      const s = doc.createElement("span");
      s.className = "g-lead";
      s.textContent = info.prefix.trim();
      lead.push(s);
    }
    for (const part of info.lead) {
      if (part.kind !== "label" || !part.text?.trim()) continue;
      const s = doc.createElement("span");
      s.className = "g-lead g-label";
      s.textContent = part.text.trim();
      lead.push(s);
    }
    row.el.replaceChildren(...lead, ta);
    row.el.classList.add("g-editing");
    setTabStop(row);
    editing = { row, ta, group: `type:${++groupCounter}` };
    ta.addEventListener("input", () => onInput(editing && editing.ta === ta ? editing : null));
    ta.addEventListener("keydown", (e) => onWordsKey(e, row, ta));
    ta.addEventListener("focusout", () => {
      if (editing && editing.ta === ta) stopEdit();
    });
    ta.focus();
    const at = caret === "start" ? 0 : caret === "end" ? ta.value.length : Math.max(0, Math.min(caret, ta.value.length));
    ta.setSelectionRange(at, at);
    grow(ta);
    placeCursor(row);
  }

  /** @param {HTMLTextAreaElement} ta */
  function grow(ta) {
    ta.style.height = "auto";
    ta.style.height = `${ta.scrollHeight}px`;
  }

  /** The words in the box are the words in the document: a change goes through the queue so none is applied out of order. @param {Editing | null} e */
  function onInput(e) {
    if (!e) return;
    // A paragraph is one line of words: line ends that were pasted in become spaces.
    if (/[\r\n]/.test(e.ta.value)) {
      const at = e.ta.selectionStart;
      e.ta.value = e.ta.value.replace(/\r\n|[\r\n]/g, " ");
      e.ta.setSelectionRange(at, at);
    }
    grow(e.ta);
    run(() => commit(e));
    tellTools();
  }

  /** @param {() => Promise<void>} job */
  function run(job) {
    queue = queue.then(job).catch((err) => {
      ui.say(messageOf(err), true);
      render();
    });
    return queue;
  }

  /** Put the words that are in the box into the document, as part of the run of typing it is in. @param {Editing} e */
  async function commit(e) {
    if (gone || !shown.includes(e.row)) return;
    const entry = entryOf(e.row);
    const xml = setWords(entry.xml ?? "", entry.style ?? "Normal", e.ta.value, io);
    if (xml === entry.xml) return;
    const at = idx(e.row);
    await guide.replaceParagraphs({ start: at, count: 1, guard: entry.text ?? "", xmls: [xml], select: at, group: e.group });
    entries = guide.exportState().paragraphs;
    entrySigs = entries.map(rowSignature);
    e.row.sig = rowSignature(entries[at]);
  }

  /**
   * @param {KeyboardEvent} ev @param {Row} row @param {HTMLTextAreaElement} ta
   */
  function onWordsKey(ev, row, ta) {
    const action = wordsKey({
      key: ev.key,
      shift: ev.shiftKey,
      ctrl: ev.ctrlKey,
      alt: ev.altKey,
      meta: ev.metaKey,
      start: ta.selectionStart,
      end: ta.selectionEnd,
      length: ta.value.length,
      composing: ev.isComposing,
    });
    if (!action) return;
    ev.preventDefault();
    const e = editing;
    if (!e || e.ta !== ta) return;
    switch (action.do) {
      case "stop":
        run(async () => {
          await commit(e);
          stopEdit();
          focusRow(row);
        });
        break;
      case "split":
        run(() => split(e));
        break;
      case "join":
        run(() => join(e));
        break;
      case "previous":
      case "next": {
        const caret = action.caret ?? "start";
        const step = action.do === "next" ? 1 : -1;
        run(async () => {
          await commit(e);
          move(row, step, caret, true);
        });
        break;
      }
      case "undo":
      case "redo":
        run(() => history(action.do === "redo", true));
        break;
    }
  }

  /**
   * Go to the row before or after: into its words when it has some the author can change and the author was in words, else onto the row.
   * @param {Row} from @param {number} step @param {"start" | "end"} caret @param {boolean} wasEditing
   */
  function move(from, step, caret, wasEditing) {
    const to = shown[shown.indexOf(from) + step];
    if (!to) return;
    if (wasEditing && to.el.dataset.editable === "true") startEdit(to, caret);
    else {
      stopEdit();
      focusRow(to);
    }
  }

  /** Enter: the words after the caret become the next paragraph (a heading just goes on to the next row). @param {Editing} e */
  async function split(e) {
    const row = e.row;
    const entry = entryOf(row);
    if (isHeading(entry.style ?? "")) {
      await commit(e);
      move(row, 1, "start", true);
      return;
    }
    const start = e.ta.selectionStart;
    const end = e.ta.selectionEnd;
    // What is selected is replaced by the Enter, as in any editor.
    if (end > start) {
      e.ta.value = e.ta.value.slice(0, start) + e.ta.value.slice(end);
      e.ta.setSelectionRange(start, start);
    }
    await commit(e);
    const now = entryOf(row);
    const cut = splitWords(now.xml ?? "", now.style ?? "Normal", start, io);
    if (!cut) {
      ui.say("This part cannot be split.");
      return;
    }
    const at = idx(row);
    // The row it was is drawn as it reads (a row whose paragraph did not change is kept, so it must not be left as a text box).
    stopEdit();
    await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls: [cut.first, cut.second], select: at + 1 });
    render();
    const next = rowAt(at + 1);
    if (next) startEdit(next, 0);
  }

  /** Backspace at the start: the paragraph joins the one before; an empty one that cannot join is removed. @param {Editing} e */
  async function join(e) {
    const row = e.row;
    await commit(e);
    const at = idx(row);
    const current = entries[at];
    const previous = entries[at - 1];
    if (!previous) return;
    const joined = joinWords(previous, current, io);
    stopEdit();
    if (joined) {
      await guide.replaceParagraphs({ start: at - 1, count: 2, guard: previous.text ?? "", xmls: [joined.xml], select: at - 1 });
      render();
      const target = rowAt(at - 1);
      if (target) startEdit(target, joined.offset);
      return;
    }
    const info = readParagraph(current.xml ?? "", current.style ?? "Normal", io);
    if (info.editable && info.text === "" && !info.lead.length) {
      // Nothing is lost by taking an empty paragraph away.
      await guide.replaceParagraphs({ start: at, count: 1, guard: current.text ?? "", xmls: [], select: at - 1 });
      render();
      const target = rowAt(at - 1);
      if (target) {
        if (target.el.dataset.editable === "true") startEdit(target, "end");
        else focusRow(target);
      }
      return;
    }
    // It stays as it is: the words go back into the box.
    render();
    const again = rowAt(at);
    if (again) startEdit(again, 0);
  }

  /** @param {boolean} redo @param {boolean} wasEditing */
  async function history(redo, wasEditing) {
    stopEdit();
    const done = redo ? guide.redo() : guide.undo();
    if (!done) {
      ui.say(redo ? "There is nothing to redo." : "There is nothing to undo.");
      if (wasEditing) {
        const r = rowNear(guide.cursor());
        if (r) startEdit(r, "end");
      }
      return;
    }
    render();
    const r = rowNear(guide.cursor());
    if (!r) return;
    if (wasEditing && r.el.dataset.editable === "true") startEdit(r, "end");
    else focusRow(r);
    ui.say(redo ? "Redone." : "Undone.");
  }

  /** @param {EventTarget | null} target @returns {Row | null} */
  function rowOf(target) {
    const el = target instanceof Element ? target.closest("[data-index]") : null;
    return el ? shown.find((r) => r.el === el) ?? null : null;
  }

  // The keys of a row that has the keyboard, and a click on one.
  page.addEventListener("keydown", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    const row = rowOf(ev.target);
    if (!row || ev.target !== row.el) return;
    const action = rowKey({ key: ev.key, shift: ev.shiftKey, ctrl: ev.ctrlKey, alt: ev.altKey, meta: ev.metaKey });
    if (!action) return;
    ev.preventDefault();
    const at = shown.indexOf(row);
    if (action === "down") focusRow(shown[at + 1] ?? row);
    else if (action === "up") focusRow(shown[at - 1] ?? row);
    else if (action === "first") focusRow(shown[0]);
    else if (action === "last") focusRow(shown[shown.length - 1]);
    else if (action === "edit") startEdit(row, "end");
    else run(() => history(action === "redo", false));
  });
  page.addEventListener("click", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    const row = rowOf(ev.target);
    if (row) startEdit(row, "end");
  });
  page.addEventListener("focusin", (ev) => {
    if (ev.target instanceof HTMLTextAreaElement) return;
    const row = rowOf(ev.target);
    if (row && ev.target === row.el) {
      setTabStop(row);
      placeCursor(row);
    }
  });

  render();
  return {
    render,
    /** The document is about to go: stop everything. */
    destroy() {
      gone = true;
      if (selectionTimer !== null) window.clearTimeout(selectionTimer);
      editing = null;
    },
    /** Put the keyboard on the row of the cursor. */
    focusCursor() {
      focusRow(rowNear(guide.cursor()));
    },
  };
}
