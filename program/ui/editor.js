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
import { KINDS, isHeading, joinWords, kindChange, nextParagraphSpec, paragraphWith, paragraphWithSpans, readParagraph, setSpans, splitWords, withStyle } from "../../src/guide-edit.mjs";
import { applyChange, cleanChange, rprAt, slice, summarize } from "../../src/guide-runs.mjs";
import { canFormat, readParagraphFormat, withParagraphFormat } from "../../src/guide-format.mjs";
import { idForName } from "../../src/guide-styles-edit.mjs";
import { addColumn, addRow, deleteColumn, deleteRow, mergeCell, readTable, setCellWords, setTableAlt, splitCell, ALT_LIMITS } from "../../src/guide-table.mjs";
import { readPicture, setPicture, DESCRIPTION_LIMIT } from "../../src/guide-picture.mjs";
import { PICTURE_WIDTHS } from "../../src/image-prepare.mjs";
import { captionLine } from "../../src/guide-reference.mjs";
import { previewStyles } from "../../src/guide-styles-edit.mjs";
import { cellKey, cellWordsKey, reconcile, rowKey, rowSignature, sameList, wordsKey, wordsLabel } from "./editor-model.js";
import { describeGuide, drawNodes } from "./preview.js";
import { uiIcon } from "./menubar.js";
import { createRichBox } from "./richbox.js";

/** @param {string} xml */
const parseXml = (xml) => new DOMParser().parseFromString(xml, "application/xml");
const io = { parse: parseXml, serialize: (/** @type {Element} */ el) => new XMLSerializer().serializeToString(el) };
/** How long after the author stops moving or typing the tools are told where the cursor is. */
const SELECTION_DELAY = 250;

/**
 * @typedef {{ xml?: string, text?: string, style?: string }} Entry
 * @typedef {{ sig: string, el: HTMLElement }} Row
 * @typedef {ReturnType<typeof createRichBox>} RichBox
 * @typedef {{ row: Row, ta: RichBox, group: string }} Editing
 */

/**
 * @param {Document} doc
 * @param {HTMLElement} host where the guide is drawn
 * @param {any} guide the guide's document, with its saving
 * @param {{ zoom?: { more: () => unknown, less: () => unknown, actual: () => unknown, level: () => number }, say: (message: string, bad?: boolean) => void, askText: (current: any, validate: (change: any) => void) => Promise<any>, openStyles: (startId?: string) => Promise<void>, askName: (options: any) => Promise<string | null>, ask?: (options: any) => Promise<any>, askFormat: (current: any, validate: (fmt: any) => void) => Promise<any>, choose: (options: { title: string, intro?: string, confirm: string, options: { id: string, label: string, note?: string }[], current?: string }) => Promise<string | null> }} ui
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
  const bar = doc.createElement("div");
  bar.className = "format-bar";
  bar.setAttribute("role", "toolbar");
  bar.setAttribute("aria-label", "Text formatting");
  host.replaceChildren(bar, page, hint);

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
  /** The rows chosen as a range (by the index of their paragraph), as with Shift+Down. @type {{ anchor: number, head: number } | null} */
  let selection = null;
  /** The row the author was last on, so a command from a menu knows which one it is for. @type {Row | null} */
  let lastRow = null;
  /** The cell of a table the author is in, and the cell whose words are being changed. @type {{ row: Row, r: number, c: number } | null} */
  let grid = null;
  /** @type {{ row: Row, r: number, c: number, ta: HTMLTextAreaElement, group: string } | null} */
  let cellEditing = null;
  let cellGroups = 0;
  /** Where words were last chosen in a paragraph, kept for a command chosen from a menu (which takes the focus from the box). @type {{ index: number, start: number, end: number, at: number } | null} */
  let textSel = null;
  /** What was last cut or copied inside the guide: whole paragraphs, to be pasted back as they were. @type {{ paragraphs: { xml: string, style: string }[], text: string } | null} */
  let clip = null;

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

  /** How the guide's styles look now (the changes the author made, drawn over the usual look), remade only when those change. */
  let stylesKey = "";
  /** @type {import("./preview.js").StyleView | undefined} */
  let styleView;
  function currentStyles() {
    const edits = guide.styleEdits();
    const key = JSON.stringify(edits);
    if (key !== stylesKey) {
      stylesKey = key;
      styleView = previewStyles(edits, guide.listStyles());
      // Every row is drawn again with the styles as they are now.
      for (const r of shown) r.sig = "";
    }
    return styleView;
  }

  /** Read the document and bring the rows up to date: only what changed is drawn again. */
  function render() {
    if (gone) return;
    entries = guide.exportState().paragraphs;
    const stylesWere = stylesKey;
    const styles = currentStyles();
    const sigs = entries.map(rowSignature);
    if (sameList(sigs, entrySigs) && stylesKey === stylesWere) return;
    entrySigs = sigs;
    const nodes = describeGuide(entries, parseXml, { blank: true, media: (id) => guide.mediaOf(id), styles });
    const rowSigs = nodes.map((n) => rowSignature(entries[n.index ?? 0]));
    const { prefix, oldEnd, newEnd } = reconcile(shown.map((r) => r.sig), rowSigs);
    // A row that is being replaced cannot go on being edited: the document changed under it (Undo, or the tools).
    if (editing && shown.indexOf(editing.row) >= prefix && shown.indexOf(editing.row) < oldEnd) editing = null;
    if (cellEditing && shown.indexOf(cellEditing.row) >= prefix && shown.indexOf(cellEditing.row) < oldEnd) cellEditing = null;
    if (grid && shown.indexOf(grid.row) >= prefix && shown.indexOf(grid.row) < oldEnd) grid = null;
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
      el.dataset.span = String(node.span ?? 1);
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
    markSelection();
    // The keyboard stays where it was, even when the row it was on has been drawn again.
    if (hadFocus && !doc.activeElement?.closest("#work-guide") && !editing) focusRow(rowNear(guide.cursor()));
  }

  /** @param {Row | null} row */
  function focusRow(row) {
    if (!row) return;
    setTabStop(row);
    lastRow = row;
    row.el.focus();
    placeCursor(row);
  }

  /** Draw a row's own words again, after they were being changed. @param {Row} row */
  function redrawRow(row) {
    const entry = entryOf(row);
    const [node] = describeGuide([entry], parseXml, { blank: true, media: (id) => guide.mediaOf(id), styles: currentStyles() });
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
    textSel = { index: idx(e.row), start: e.ta.selectionStart, end: e.ta.selectionEnd, at: Date.now() };
    redrawRow(e.row);
  }

  /**
   * Change the words of a row.
   * @param {Row} row @param {number | "start" | "end"} caret
   */
  function startEdit(row, caret) {
    if (row.el.tagName === "TABLE") {
      stopEdit();
      focusCell(row, 0, 0);
      return;
    }
    stopCellEdit();
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
    // Words chosen in another paragraph are not these.
    if (textSel && textSel.index !== idx(row)) textSel = null;
    const ta = createRichBox(doc, { spans: info.spans, label: wordsLabel(entry.style ?? "Normal"), describedby: "edit-hint", io, styleCss: (id) => currentStyles()?.cssForId?.(id) ?? null });
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
    // The picture of a block's kind stays at the front while its words are changed.
    const icon = row.el.querySelector(":scope > img.g-icon");
    if (icon) lead.unshift(/** @type {HTMLElement} */ (icon.cloneNode(true)));
    row.el.replaceChildren(...lead, ta.el);
    row.el.classList.add("g-editing");
    setTabStop(row);
    editing = { row, ta, group: `type:${++groupCounter}` };
    lastRow = row;
    clearSelection();
    ta.el.addEventListener("input", () => onInput(editing && editing.ta === ta ? editing : null));
    ta.el.addEventListener("keydown", (e) => onWordsKey(e, row, ta));
    ta.el.addEventListener("paste", (e) => onWordsPaste(e, row, ta));
    ta.el.addEventListener("focusout", () => {
      if (editing && editing.ta === ta) stopEdit();
    });
    ta.focus();
    const at = caret === "start" ? 0 : caret === "end" ? ta.value.length : Math.max(0, Math.min(caret, ta.value.length));
    ta.setSelectionRange(at, at);
    placeCursor(row);
  }

  /** A plain text box (a cell of a table) grows with its words; the box for a paragraph's words does by itself. @param {HTMLTextAreaElement} ta */
  function grow(ta) {
    ta.style.height = "auto";
    ta.style.height = `${ta.scrollHeight}px`;
  }

  /** The words in the box are the words in the document: a change goes through the queue so none is applied out of order. @param {Editing | null} e */
  function onInput(e) {
    if (!e) return;
    const group = e.group;
    run(() => commit(e, group));
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

  /** Put the words that are in the box into the document, as part of the run of typing it is in (or of the group given). @param {Editing} e @param {string} [group] */
  async function commit(e, group = e.group) {
    if (gone || !shown.includes(e.row)) return;
    const entry = entryOf(e.row);
    const xml = setSpans(entry.xml ?? "", entry.style ?? "Normal", e.ta.spans(), io);
    if (xml === entry.xml) return;
    const at = idx(e.row);
    await guide.replaceParagraphs({ start: at, count: 1, guard: entry.text ?? "", xmls: [xml], select: at, group });
    entries = guide.exportState().paragraphs;
    entrySigs = entries.map(rowSignature);
    e.row.sig = rowSignature(entries[at]);
  }

  /**
   * @param {KeyboardEvent} ev @param {Row} row @param {RichBox} ta
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
      case "align": {
        const at = idx(row);
        const caret = ta.selectionStart;
        run(async () => {
          await commit(e);
          await formatRows({ alignment: action.value ?? null });
          const again = rowAt(at);
          if (again) startEdit(again, caret);
        });
        break;
      }
      case "undo":
      case "redo":
        run(() => history(action.do === "redo", true));
        break;
      case "text":
        run(() => textFormat(action.setting ?? "bold"));
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
    if (end > start) e.ta.setRangeText("", start, end, "preserve");
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

  /**
   * Put these settings on the paragraphs the author is on or has chosen (a value left out stays, null lets the style decide again). One step of Undo.
   * @param {Record<string, unknown>} fmt
   */
  async function formatRows(fmt) {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    stopEdit();
    stopCellEdit();
    const group = `format:${++groupCounter}`;
    let changed = 0;
    let skipped = "";
    for (const row of rows) {
      const entry = entryOf(row);
      const can = canFormat(spanOf(row) === 1 ? entry.xml ?? "" : "", entry.style ?? "Normal", io);
      if (!can.ok) {
        skipped = can.reason ?? "";
        continue;
      }
      let xml;
      try {
        xml = withParagraphFormat(entry.xml ?? "", entry.style ?? "Normal", fmt, io);
      } catch (e) {
        ui.say(messageOf(e), true);
        return;
      }
      if (xml === entry.xml) continue;
      const at = idx(row);
      await guide.replaceParagraphs({ start: at, count: 1, guard: entry.text ?? "", xmls: [xml], select: at, group });
      entries = guide.exportState().paragraphs;
      changed++;
    }
    entrySigs = [];
    render();
    if (changed) ui.say(changed === 1 ? "Formatted. Ctrl+Z takes it back." : `${changed} paragraphs formatted. Ctrl+Z takes it back.`);
    else if (skipped) ui.say(skipped);
    focusRow(rowNear(idx(rows[rows.length - 1])));
  }

  // ---- the look of words inside a paragraph

  /** The settings the paragraph's style gives its words (whether a toggle turns a look on or off depends on them). @param {string} styleName @returns {Record<string, any>} */
  function baseFont(styleName) {
    return guide.listStyles().find((/** @type {any} */ s) => s.name === styleName)?.effective?.font ?? {};
  }

  /**
   * The change a command makes, from what the words are like now: a look that is on goes off (explicitly, when the style has it on), a look that is
   * not on, or only on some of the words, goes on. A change that came from the dialog is used as it is.
   * @param {string | Record<string, unknown>} what @param {Record<string, any>} summary @param {Record<string, any>} base
   * @returns {Record<string, unknown>}
   */
  function changeFor(what, summary, base) {
    if (typeof what !== "string") return what;
    if (what === "clear") return { bold: null, italic: null, underline: null, strike: null, script: null, color: null, font: null, size: null, style: null };
    if (what === "superscript" || what === "subscript") return { script: summary.script === what ? null : what };
    const on = summary[what] === true || (summary[what] === undefined && base[what] === true);
    if (on) return { [what]: base[what] === true ? false : null };
    return { [what]: base[what] === true ? null : true };
  }

  /** A few words for what a change does, for the author. @param {Record<string, unknown>} change */
  function sayChange(change) {
    const names = { bold: "bold", italic: "italic", underline: "underlined", strike: "struck through", script: "raised or lowered", color: "colored", font: "in another font", size: "another size", style: "in a character style" };
    const parts = Object.entries(change).map(([k, v]) => (v === null || v === false ? `no longer ${/** @type {Record<string, string>} */ (names)[k] ?? k}` : /** @type {Record<string, string>} */ (names)[k] ?? k));
    return parts.length === 1 ? parts[0] : "given a new look";
  }

  /**
   * Give words a look. While the words are being changed, the ones chosen (with none chosen, the next ones typed); from a menu, the words that
   * were chosen a moment ago, else every word of the paragraph or paragraphs the author is on. One step of Undo.
   * @param {string | Record<string, unknown>} what a setting to switch ("bold"...), "clear", or a change from the dialog
   * @param {{ index: number, start: number, end: number } | null} [picked] words chosen before a dialog took the focus
   */
  async function textFormat(what, picked = null) {
    const e = editing;
    if (e && !picked) {
      const style = entryOf(e.row).style ?? "Normal";
      let change;
      try {
        change = cleanChange(changeFor(what, e.ta.state(), baseFont(style)));
      } catch (err) {
        ui.say(messageOf(err), true);
        return;
      }
      // The change is one step of Undo of its own, and typing after it is another.
      e.group = `text:${++groupCounter}`;
      const changed = e.ta.apply(change);
      e.group = `type:${++groupCounter}`;
      if (!changed) ui.say(`The next words you type will be ${sayChange(change)}.`);
      syncBar();
      return;
    }
    const rows = picked ? [rowAt(picked.index)].filter((r) => r !== null) : targetRows();
    if (!rows.length) return;
    const recent = picked ?? (textSel && rows.length === 1 && idx(rows[0]) === textSel.index && textSel.end > textSel.start && Date.now() - textSel.at < 20000 ? textSel : null);
    textSel = null;
    await settle();
    stopEdit();
    stopCellEdit();
    const group = `text:${++groupCounter}`;
    let changed = 0;
    let skipped = "";
    /** @type {Record<string, unknown> | null} */
    let change = null;
    for (const row of /** @type {Row[]} */ (rows)) {
      const entry = entryOf(row);
      const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
      if (!info.editable) {
        skipped = info.reason ?? "";
        continue;
      }
      const from = recent ? recent.start : 0;
      const to = recent ? recent.end : info.text.length;
      try {
        // What to do is decided once, from the first paragraph, and the same is done to the rest.
        change ??= cleanChange(changeFor(what, summarize(info.spans, from, to, io), baseFont(entry.style ?? "Normal")));
        const xml = setSpans(entry.xml ?? "", entry.style ?? "Normal", applyChange(info.spans, from, to, /** @type {any} */ (change), io), io);
        if (xml === entry.xml) continue;
        const at = idx(row);
        await guide.replaceParagraphs({ start: at, count: 1, guard: entry.text ?? "", xmls: [xml], select: at, group });
        entries = guide.exportState().paragraphs;
        changed++;
      } catch (err) {
        ui.say(messageOf(err), true);
        return;
      }
    }
    entrySigs = [];
    render();
    if (changed) ui.say(`${changed === 1 ? "The words are" : `${changed} paragraphs are`} ${sayChange(change ?? {})}. Ctrl+Z takes it back.`);
    else if (skipped) ui.say(skipped);
    const back = rowNear(idx(/** @type {Row[]} */ (rows)[rows.length - 1]));
    if (recent && rows.length === 1 && back && back.el.dataset.editable === "true") {
      // Back into the words, with the same ones chosen.
      startEdit(back, recent.start);
      editing?.ta.setSelectionRange(recent.start, recent.end);
    } else focusRow(back);
  }

  /** The words a command from a menu or a dialog is for: the ones chosen (while changing the words, or a moment ago), else the paragraph the author is on. */
  function chosenWords() {
    const recent = textSel && textSel.end > textSel.start && Date.now() - textSel.at < 20000 && rowAt(textSel.index) ? textSel : null;
    const sel = editing ? { index: idx(editing.row), start: editing.ta.selectionStart, end: editing.ta.selectionEnd } : recent;
    const rows = sel ? [rowAt(sel.index)].filter((r) => r !== null) : targetRows();
    return { sel, row: /** @type {Row | undefined} */ (rows[0]) };
  }

  /** Go back into the words, with the same ones chosen. @param {Row} row @param {{ start: number, end: number } | null} sel */
  function backInto(row, sel) {
    if (!sel) return focusRow(row);
    startEdit(row, sel.start);
    editing?.ta.setSelectionRange(sel.start, sel.end);
  }

  /** Format, Character style...: give the chosen words one of the character styles the author made, or none. */
  async function charStyleDialog() {
    const { sel, row } = chosenWords();
    if (!row) return;
    const mine = guide.listStyles().filter((/** @type {any} */ s) => s.type === "character" && s.custom);
    if (!mine.length) return ui.say("You have no character styles yet. Choose Format, Styles..., then New character style, and come back.");
    await settle();
    const entry = entryOf(row);
    const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!info.editable) return ui.say(info.reason ?? "These words are changed with the tools beside the guide.");
    const now = summarize(info.spans, sel ? sel.start : 0, sel ? sel.end : info.text.length, io);
    const chosen = await ui.choose({
      title: "Character style",
      intro: "A character style gives the chosen words (with none chosen, the whole paragraph) a named look, which you can change in Styles... for every word that has it.",
      confirm: "Apply",
      options: [{ id: "-", label: "No character style", note: "The paragraph's style decides." }, ...mine.map((/** @type {any} */ s) => ({ id: s.id, label: s.name }))],
      current: typeof now.style === "string" ? now.style : "-",
    });
    if (chosen === null) return backInto(row, sel);
    await textFormat({ style: chosen === "-" ? null : chosen }, sel ?? { index: idx(row), start: 0, end: info.text.length });
  }

  /** Format, Text...: the dialog for the look of the words that are chosen (or of the whole paragraph). */
  async function textDialog() {
    const { sel, row } = chosenWords();
    if (!row) return;
    await settle();
    const entry = entryOf(row);
    const info = readParagraph(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!info.editable) return ui.say(info.reason ?? "These words are changed with the tools beside the guide.");
    const base = baseFont(entry.style ?? "Normal");
    const from = sel ? sel.start : 0;
    const to = sel ? sel.end : info.text.length;
    const now = summarize(info.spans, from, to, io);
    /** @type {Record<string, any>} */
    const current = { font: typeof now.font === "string" ? now.font : base.name, size: typeof now.size === "number" ? now.size : base.size, color: typeof now.color === "string" ? now.color : base.color, script: now.script === "mixed" ? undefined : now.script };
    for (const k of ["bold", "italic", "underline", "strike"]) current[k] = now[k] === true || (now[k] === undefined && base[k] === true);
    const asked = await ui.askText(current, (change) => cleanChange(change));
    if (!asked) return backInto(row, sel);
    if (!Object.keys(asked).length) {
      ui.say("Nothing was changed.");
      return backInto(row, sel);
    }
    // The dialog gives the looks that were switched as on or off; what "off" means depends on the style.
    /** @type {Record<string, unknown>} */
    const change = { ...asked };
    for (const k of ["bold", "italic", "underline", "strike"]) if (k in change) change[k] = change[k] === true ? (base[k] === true ? null : true) : base[k] === true ? false : null;
    await textFormat(change, sel ?? { index: idx(row), start: 0, end: info.text.length });
  }

  /** The Styles panel, on the style of the paragraph the author is on. */
  async function stylesPanel() {
    const row = targetRows()[0] ?? null;
    await settle();
    stopEdit();
    const name = row ? entryOf(row).style : undefined;
    const startId = name ? guide.listStyles().find((/** @type {any} */ s) => s.name === name)?.id : undefined;
    await ui.openStyles(startId);
    entrySigs = [];
    stylesKey = "";
    render();
    // The keyboard goes back to the row it was on, unless the author has already gone on into the guide.
    if (!doc.activeElement?.closest("#work-guide")) focusRow(row && shown.includes(row) ? row : rowNear(guide.cursor()));
  }

  /** A style made from the paragraph the author is on: its alignment and spacing move into the style, and the paragraph takes the style. */
  async function newStyleFromParagraph() {
    const rows = targetRows();
    if (!rows.length) return;
    const row = rows[0];
    await settle();
    stopEdit();
    const entry = entryOf(row);
    const can = kindChange(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!can.ok || spanOf(row) !== 1) return ui.say(can.reason ?? "A style cannot be made from this.");
    const name = await ui.askName({ title: "New style from this paragraph", label: "Name of the style", confirm: "Create", hint: "The paragraph's alignment, spacing and indents move into the new style, and the paragraph takes the style." });
    if (!name) {
      focusRow(row);
      return;
    }
    const clean = name.replace(/\s+/g, " ").trim();
    const list = guide.listStyles();
    if (list.some((/** @type {any} */ s) => s.name.toLowerCase() === clean.toLowerCase())) return ui.say(`There is already a style called ${clean}.`, true);
    const base = list.find((/** @type {any} */ s) => s.name === entry.style);
    if (!base) return ui.say("The style this paragraph is in could not be found.", true);
    const fmt = readParagraphFormat(entry.xml ?? "", io);
    /** @type {Record<string, unknown>} */
    const paragraph = {};
    for (const [k, v] of Object.entries(fmt)) if (v !== null) paragraph[k] = v;
    const edits = guide.styleEdits();
    const id = idForName(clean, [...list.map((/** @type {any} */ s) => s.id), ...edits.custom.map((/** @type {any} */ c) => c.id)]);
    edits.custom.push({ id, name: clean, type: "paragraph", basedOn: base.id, spec: Object.keys(paragraph).length ? { paragraph } : {} });
    const group = `newstyle:${++groupCounter}`;
    try {
      await guide.setStyleEdits(edits, { group });
      const at = idx(row);
      const now = entryOf(row);
      const moved = withStyle(now.xml ?? "", now.style ?? "Normal", id, io);
      const cleared = withParagraphFormat(moved, clean, { alignment: null, spaceBefore: null, spaceAfter: null, lineSpacing: null, leftIndent: null, firstLineIndent: null }, io);
      await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls: [cleared], select: at, group });
    } catch (e) {
      ui.say(messageOf(e), true);
      return;
    }
    entries = guide.exportState().paragraphs;
    entrySigs = [];
    render();
    focusRow(rowNear(idx(row)));
    ui.say(`${clean} is a style of your own, and this paragraph is in it. Ctrl+Z takes it back.`);
  }

  async function paragraphDialog() {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const first = entryOf(rows[0]);
    const can = canFormat(spanOf(rows[0]) === 1 ? first.xml ?? "" : "", first.style ?? "Normal", io);
    if (!can.ok) return ui.say(can.reason ?? "This cannot be formatted.");
    stopEdit();
    const current = readParagraphFormat(first.xml ?? "", io);
    const fmt = await ui.askFormat(current, (f) => void withParagraphFormat(first.xml ?? "", first.style ?? "Normal", f, io));
    if (!fmt) {
      focusRow(rows[0]);
      return;
    }
    await formatRows(fmt);
  }

  /** @param {boolean} redo @param {boolean} wasEditing */
  async function history(redo, wasEditing) {
    stopEdit();
    // The paragraphs that were chosen may not be the same paragraphs after this: nothing stays chosen.
    selection = null;
    stopCellEdit();
    grid = null;
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

  // ---- rows chosen together, and what is done to them

  /** @param {Row} row How many entries of the document the row covers: a table covers its cells. */
  const spanOf = (row) => Number(row.el.dataset.span || 1);

  function selectedRows() {
    if (!selection) return [];
    const lo = Math.min(selection.anchor, selection.head);
    const hi = Math.max(selection.anchor, selection.head);
    return shown.filter((r) => idx(r) >= lo && idx(r) <= hi);
  }

  function markSelection() {
    const chosen = new Set(selectedRows());
    if (selection && !chosen.size) selection = null;
    for (const r of shown) {
      if (chosen.has(r)) r.el.dataset.selected = "true";
      else delete r.el.dataset.selected;
    }
  }

  function clearSelection() {
    if (!selection) return;
    selection = null;
    markSelection();
  }

  /** @param {number} anchor @param {number} head */
  function select(anchor, head) {
    selection = { anchor, head };
    markSelection();
    const n = selectedRows().length;
    ui.say(`${n} paragraph${n === 1 ? "" : "s"} selected.`);
  }

  /** Shift+Down and Shift+Up: the range grows or shrinks by a row, and the keyboard goes with it. @param {Row} row @param {number} step */
  function extend(row, step) {
    const to = shown[shown.indexOf(row) + step];
    if (!to) return;
    select(selection ? selection.anchor : idx(row), idx(to));
    focusRow(to);
  }

  /** The rows a command is for: the ones chosen together, or the one the author is on. */
  /** Whether the paragraph the author is on holds one picture of theirs, which the Edit menu's Picture details can change. */
  function onPicture() {
    const rows = targetRows();
    return rows.length === 1 && !!readPicture(entryOf(rows[0]).xml ?? "", io);
  }

  function targetRows() {
    const chosen = selectedRows();
    if (chosen.length) return chosen;
    const one = editing?.row ?? (lastRow && shown.includes(lastRow) ? lastRow : null);
    return one ? [one] : [];
  }

  async function settle() {
    if (editing) await commit(editing);
    if (cellEditing) await commitCell(cellEditing);
  }

  /** @param {string} name */
  async function command(name) {
    switch (name) {
      case "delete":
        return remove();
      case "moveUp":
        return move1(-1);
      case "moveDown":
        return move1(1);
      case "add":
        return addBelow();
      case "kind":
        return changeKind();
      case "copy":
        return copy(false);
      case "cut":
        return copy(true);
      case "paste":
        return paste(null);
      case "alignLeft":
      case "alignCenter":
      case "alignRight":
      case "alignJustify":
        return formatRows({ alignment: /** @type {"left" | "center" | "right" | "justify"} */ (name.slice(5).toLowerCase()) });
      case "clearFormat":
        return formatRows({ alignment: null, spaceBefore: null, spaceAfter: null, lineSpacing: null, leftIndent: null, firstLineIndent: null });
      case "paragraph":
        return paragraphDialog();
      case "bold":
      case "italic":
      case "underline":
      case "strike":
      case "superscript":
      case "subscript":
        return textFormat(name);
      case "clearText":
        return textFormat("clear");
      case "text":
        return textDialog();
      case "charStyle":
        return charStyleDialog();
      case "styles":
        return stylesPanel();
      case "newStyle":
        return newStyleFromParagraph();
      case "addRow":
        return grid ? addTableRow(grid.row, grid.r, grid.c, false) : ui.say("Move into a table first: Enter on a table starts its cells.");
      case "deleteRow":
        return grid ? deleteTableRow(grid.row, grid.r, grid.c) : ui.say("Move into a table first: Enter on a table starts its cells.");
      case "addColumn":
      case "deleteColumn":
      case "mergeCell":
      case "splitCell":
      case "tableDetails":
        return grid ? tableChange(name, grid.row, grid.r, grid.c) : ui.say("Move into a table first: Enter on a table starts its cells.");
      case "pictureDetails":
        return pictureDetails();
      case "undo":
        return history(false, false);
      case "redo":
        return history(true, false);
    }
  }

  async function remove() {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const start = idx(rows[0]);
    const last = rows[rows.length - 1];
    const end = idx(last) + spanOf(last) - 1;
    stopEdit();
    await guide.deleteRange({ start, end, guard: entries[start]?.text ?? "" });
    selection = null;
    render();
    focusRow(rowNear(Math.max(0, start - 1)) ?? shown[0] ?? null);
    ui.say(rows.length === 1 ? "Paragraph deleted. Ctrl+Z brings it back." : `${rows.length} paragraphs deleted. Ctrl+Z brings them back.`);
  }

  /** @param {number} step */
  async function move1(step) {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const first = rows[0];
    const last = rows[rows.length - 1];
    const start = idx(first);
    const end = idx(last) + spanOf(last) - 1;
    let after;
    if (step < 0) {
      const above = shown[shown.indexOf(first) - 1];
      if (!above) return ui.say("It is already at the top.");
      after = idx(above) - 1;
    } else {
      const below = shown[shown.indexOf(last) + 1];
      if (!below) return ui.say("It is already at the end.");
      after = idx(below) + spanOf(below) - 1;
    }
    stopEdit();
    const moved = await guide.moveRange({ start, end, guard: entries[start]?.text ?? "", after });
    render();
    select(moved.at, moved.at + (end - start));
    focusRow(rowNear(moved.at));
  }

  async function addBelow() {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    const last = rows[rows.length - 1];
    const entry = entryOf(last);
    const at = idx(last) + spanOf(last);
    const spec = spanOf(last) === 1 ? nextParagraphSpec(entry.xml ?? "", entry.style ?? "Normal", io) : null;
    const xml = paragraphWith(spec ?? { styleId: null, numPr: null }, "", io);
    stopEdit();
    await guide.replaceParagraphs({ start: at, count: 0, guard: "", xmls: [xml], select: at });
    selection = null;
    render();
    const row = rowAt(at);
    if (row) startEdit(row, 0);
  }

  async function changeKind() {
    const rows = targetRows();
    if (rows.length !== 1) return ui.say("Choose one paragraph to change its kind.");
    const row = rows[0];
    await settle();
    const entry = entryOf(row);
    const can = kindChange(entry.xml ?? "", entry.style ?? "Normal", io);
    if (!can.ok) return ui.say(can.reason ?? "This kind cannot be changed.");
    stopEdit();
    // The kinds the program offers, then the author's own styles.
    const mine = guide.styleEdits().custom.filter((/** @type {any} */ c) => c.type === "paragraph");
    const kinds = [
      ...KINDS.map((k) => ({ id: k.id, label: k.label, note: k.note, style: k.style, styleId: /** @type {string | null} */ (k.styleId) })),
      ...mine.map((/** @type {any} */ c) => ({ id: `style:${c.id}`, label: c.name, note: "A style of your own.", style: c.name, styleId: /** @type {string | null} */ (c.id) })),
    ];
    const chosen = await ui.choose({
      title: "Kind of paragraph",
      intro: "What this paragraph is. The tools and the participant workbook treat each kind in its own way.",
      confirm: "Change",
      options: kinds.map((k) => ({ id: k.id, label: k.label, note: k.note })),
      current: kinds.find((k) => k.style === entry.style)?.id,
    });
    const kind = kinds.find((k) => k.id === chosen);
    if (!kind || kind.style === entry.style) {
      focusRow(row);
      return;
    }
    const at = idx(row);
    const now = entryOf(row);
    const xml = withStyle(now.xml ?? "", now.style ?? "Normal", kind.styleId, io);
    await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls: [xml], select: at });
    render();
    focusRow(rowAt(at));
    ui.say(`Now ${kind.label.toLowerCase()}.`);
  }

  /** @param {boolean} cut */
  async function copy(cut) {
    const rows = targetRows();
    if (!rows.length) return;
    await settle();
    /** @type {{ xml: string, style: string }[]} */
    const paragraphs = [];
    for (const r of rows) {
      const e = entryOf(r);
      if (spanOf(r) !== 1 || !/^<w:p[\s>]/.test(e.xml ?? "")) return ui.say("Tables are moved and deleted, not copied, in this version.");
      paragraphs.push({ xml: e.xml ?? "", style: e.style ?? "Normal" });
    }
    const text = rows.map((r) => (entryOf(r).text ?? "").trim()).join("\n");
    clip = { paragraphs, text };
    // The system clipboard is a courtesy (it may refuse): the guide keeps its own copy.
    navigator.clipboard?.writeText(text).catch(() => {});
    if (cut) await remove();
    else ui.say(rows.length === 1 ? "Copied." : `${rows.length} paragraphs copied.`);
  }

  /** Paste after the rows the author is on. @param {string | null} text what the system clipboard held, or null from a menu */
  async function paste(text) {
    const rows = targetRows();
    const last = rows[rows.length - 1] ?? shown[shown.length - 1];
    if (!last) return;
    await settle();
    const plain = (text ?? "").replace(/\r\n?/g, "\n").trim();
    /** @type {string[]} */
    let xmls = [];
    if (clip && (text === null || plain === "" || plain === clip.text.trim())) xmls = clip.paragraphs.map((p) => p.xml);
    else if (plain) {
      const entry = entryOf(last);
      const spec = (spanOf(last) === 1 ? nextParagraphSpec(entry.xml ?? "", entry.style ?? "Normal", io) : null) ?? { styleId: null, numPr: null };
      xmls = plain
        .split("\n")
        .filter((l) => l.trim())
        .map((l) => paragraphWith(spec, l, io));
    }
    if (!xmls.length) return ui.say("There is nothing to paste. Copy or cut some paragraphs first.");
    const at = idx(last) + spanOf(last);
    stopEdit();
    await guide.replaceParagraphs({ start: at, count: 0, guard: "", xmls, select: at });
    render();
    select(at, at + xmls.length - 1);
    focusRow(rowNear(at));
  }

  /**
   * Several lines pasted into a paragraph's words: the first goes in at the caret, the others become paragraphs after it, the end of the
   * last one is the text that was after the caret. In a heading, which is one line, they are joined with spaces.
   * @param {ClipboardEvent} ev @param {Row} row @param {RichBox} ta
   */
  function onWordsPaste(ev, row, ta) {
    const raw = ev.clipboardData?.getData("text/plain") ?? "";
    const lines = raw.split(/\r\n|[\r\n]/).filter((l) => l.trim());
    if (lines.length < 2) return;
    ev.preventDefault();
    const e = editing;
    if (!e || e.ta !== ta) return;
    const style = entryOf(row).style ?? "Normal";
    if (isHeading(style)) {
      ta.setRangeText(lines.join(" "), ta.selectionStart, ta.selectionEnd, "end");
      return;
    }
    run(async () => {
      await commit(e);
      const now = entryOf(row);
      const spec = nextParagraphSpec(now.xml ?? "", now.style ?? "Normal", io);
      if (!spec) return;
      const spans = ta.spans();
      const from = ta.selectionStart;
      const to = ta.selectionEnd;
      // The pasted words take the look of the character before the caret; the words after the selection keep their own.
      const look = rprAt(spans, from);
      const xmls = [
        setSpans(now.xml ?? "", now.style ?? "Normal", [...slice(spans, 0, from), { text: lines[0], rpr: look }], io),
        ...lines.slice(1, -1).map((l) => paragraphWithSpans(spec, [{ text: l, rpr: look }], io)),
        paragraphWithSpans(spec, [{ text: lines[lines.length - 1], rpr: look }, ...slice(spans, to)], io),
      ];
      const at = idx(row);
      stopEdit();
      await guide.replaceParagraphs({ start: at, count: 1, guard: now.text ?? "", xmls, select: at + xmls.length - 1 });
      render();
      const last = rowAt(at + xmls.length - 1);
      if (last) startEdit(last, lines[lines.length - 1].length);
    });
  }

  /**
   * Put a block (a form, a table, lines to write on) after the row the author is on, choose what was put in, and say how to take it back.
   * @param {(ctx: any) => string} build @param {string} what
   */
  async function insertBlock(build, what) {
    const rows = targetRows();
    const last = rows[rows.length - 1] ?? shown[shown.length - 1] ?? null;
    await settle();
    stopEdit();
    const after = last ? idx(last) + spanOf(last) - 1 : -1;
    const before = entries.length;
    await guide.insert(build, last ? { where: "after", after, spare: false, select: "last" } : { where: "start", spare: false, select: "last" });
    selection = null;
    render();
    const added = entries.length - before;
    if (added > 0) {
      select(after + 1, after + added);
      focusRow(rowNear(after + 1));
    }
    ui.say(`${what} inserted. Ctrl+Z takes it back.`);
  }

  // ---- the cells of a table: the arrow keys move between cells, Enter changes a cell's words, Tab goes on to the next cell

  /** @param {Row} row */
  const tableOf = (row) => readTable(entryOf(row).xml ?? "", io);
  /** @param {Row} row @param {number} r @param {number} c */
  const cellEl = (row, r, c) => /** @type {HTMLElement | null} */ (row.el.querySelector(`[data-row="${r}"][data-col="${c}"]`));

  /** Put the keyboard on a cell of a table, or the nearest one that is there. @param {Row} row @param {number} r @param {number} c */
  function focusCell(row, r, c) {
    const t = tableOf(row);
    const rr = Math.max(0, Math.min(r, t.rows.length - 1));
    const cc = Math.max(0, Math.min(c, (t.rows[rr]?.cells.length ?? 1) - 1));
    const el = cellEl(row, rr, cc);
    if (!el) return;
    for (const other of row.el.querySelectorAll("[data-row]")) /** @type {HTMLElement} */ (other).tabIndex = -1;
    el.tabIndex = 0;
    grid = { row, r: rr, c: cc };
    lastRow = row;
    el.focus();
    placeCursor(row);
  }

  /** @param {KeyboardEvent} ev @param {HTMLElement} el */
  function onCellKey(ev, el) {
    const row = rowOf(el);
    if (!row) return;
    const r = Number(el.dataset.row);
    const c = Number(el.dataset.col);
    const action = cellKey({ key: ev.key, shift: ev.shiftKey, ctrl: ev.ctrlKey, alt: ev.altKey, meta: ev.metaKey });
    if (!action) return;
    ev.preventDefault();
    const t = tableOf(row);
    if (action === "left") focusCell(row, r, c - 1);
    else if (action === "right") focusCell(row, r, c + 1);
    else if (action === "up") focusCell(row, r - 1, c);
    else if (action === "down") focusCell(row, r + 1, c);
    else if (action === "rowStart") focusCell(row, r, 0);
    else if (action === "rowEnd") focusCell(row, r, 999);
    else if (action === "top") focusCell(row, 0, 0);
    else if (action === "bottom") focusCell(row, t.rows.length - 1, 0);
    else if (action === "edit") startCellEdit(row, r, c, "end");
    else if (action === "leave") {
      grid = null;
      focusRow(row);
    } else if (action === "addRow") run(() => addTableRow(row, r, c, false));
    else if (action === "deleteRow") run(() => deleteTableRow(row, r, c));
    else run(() => history(action === "redo", false));
  }

  /** @param {Row} row @param {number} r @param {number} c @param {number | "start" | "end"} caret */
  function startCellEdit(row, r, c, caret) {
    const t = tableOf(row);
    const cell = t.rows[r]?.cells[c];
    if (!cell) return;
    if (!cell.editable) {
      focusCell(row, r, c);
      ui.say(cell.reason ?? "This cell is changed with the tools beside the guide.");
      return;
    }
    stopEdit();
    stopCellEdit();
    const el = cellEl(row, r, c);
    if (!el) return;
    const ta = doc.createElement("textarea");
    ta.className = "g-edit";
    ta.rows = 1;
    ta.value = cell.text;
    ta.spellcheck = true;
    const column = t.rows[0]?.cells[c]?.text.trim() || `column ${c + 1}`;
    ta.setAttribute("aria-label", r === 0 ? `Column name ${c + 1}` : `${column}, row ${r}`);
    el.replaceChildren(ta);
    /** @type {{ row: Row, r: number, c: number, ta: HTMLTextAreaElement, group: string }} */
    const ce = { row, r, c, ta, group: `cell:${++cellGroups}` };
    cellEditing = ce;
    grid = { row, r, c };
    lastRow = row;
    ta.addEventListener("input", () => {
      if (/[\r\n]/.test(ta.value)) {
        const at = ta.selectionStart;
        ta.value = ta.value.replace(/\r\n|[\r\n]/g, " ");
        ta.setSelectionRange(at, at);
      }
      grow(ta);
      run(() => commitCell(ce));
      tellTools();
    });
    ta.addEventListener("keydown", (e) => onCellWordsKey(e, ce));
    ta.addEventListener("focusout", () => {
      if (cellEditing === ce) stopCellEdit();
    });
    ta.focus();
    const at = caret === "start" ? 0 : caret === "end" ? ta.value.length : Math.max(0, Math.min(caret, ta.value.length));
    ta.setSelectionRange(at, at);
    grow(ta);
    placeCursor(row);
  }

  /** Put the words of the cell that is being changed into the table in the document. @param {{ row: Row, r: number, c: number, ta: HTMLTextAreaElement, group: string }} ce */
  async function commitCell(ce) {
    if (gone || !shown.includes(ce.row)) return;
    const entry = entryOf(ce.row);
    const xml = setCellWords(entry.xml ?? "", ce.r, ce.c, ce.ta.value, io);
    if (xml === entry.xml) return;
    const at = idx(ce.row);
    await guide.replaceTable({ start: at, guard: entry.text ?? "", xml, select: at, group: ce.group });
    entries = guide.exportState().paragraphs;
    entrySigs = entries.map(rowSignature);
    ce.row.sig = rowSignature(entries[at]);
  }

  function stopCellEdit() {
    if (!cellEditing) return;
    const ce = cellEditing;
    cellEditing = null;
    const el = shown.includes(ce.row) ? cellEl(ce.row, ce.r, ce.c) : null;
    if (!el) return;
    const cell = tableOf(ce.row).rows[ce.r]?.cells[ce.c];
    el.replaceChildren(doc.createTextNode(cell?.text ?? ""));
  }

  /**
   * Keys in the words of a cell: Enter goes down a row, Tab on to the next cell that can be written in (and, from the last one, to a new row),
   * the arrow keys at the ends of the words go to the next cell.
   * @param {KeyboardEvent} ev @param {{ row: Row, r: number, c: number, ta: HTMLTextAreaElement, group: string }} ce
   */
  function onCellWordsKey(ev, ce) {
    const ta = ce.ta;
    const action = cellWordsKey({
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
    if (cellEditing !== ce) return;
    const { row, r, c } = ce;
    run(async () => {
      await commitCell(ce);
      if (action.do === "undo" || action.do === "redo") return history(action.do === "redo", false);
      if (action.do === "stop") {
        stopCellEdit();
        focusCell(row, r, c);
        return;
      }
      const t = tableOf(row);
      /** @param {number} rr @param {number} cc @param {"start" | "end"} caret */
      const go = (rr, cc, caret) => {
        if (rr < 0 || rr >= t.rows.length || cc < 0 || cc >= t.rows[rr].cells.length) return;
        if (t.rows[rr].cells[cc].editable) startCellEdit(row, rr, cc, caret);
        else {
          stopCellEdit();
          focusCell(row, rr, cc);
        }
      };
      if (action.do === "down") go(r + 1, c, "end");
      else if (action.do === "left") go(c > 0 ? r : r - 1, c > 0 ? c - 1 : (t.rows[r - 1]?.cells.length ?? 1) - 1, "end");
      else if (action.do === "right") go(c + 1 < t.rows[r].cells.length ? r : r + 1, c + 1 < t.rows[r].cells.length ? c + 1 : 0, "start");
      else if (action.do === "up") go(r - 1, c, "end");
      else if (action.do === "downCell") go(r + 1, c, "start");
      else {
        // Tab and Shift+Tab: the next or the last cell that can be written in, in reading order.
        const step = action.do === "next" ? 1 : -1;
        /** @type {[number, number][]} */
        const order = [];
        t.rows.forEach((rowData, rr) => rowData.cells.forEach((cell, cc) => cell.editable && order.push([rr, cc])));
        const here = order.findIndex(([rr, cc]) => rr === r && cc === c);
        const to = order[here + step];
        if (to) go(to[0], to[1], step > 0 ? "start" : "end");
        else if (step > 0) await addTableRow(row, r, c, true);
      }
    });
  }

  /** @param {Row} row @param {number} r @param {number} c @param {boolean} toFirst go into the first cell of the new row that can be written in */
  async function addTableRow(row, r, c, toFirst) {
    await settle();
    stopCellEdit();
    const entry = entryOf(row);
    const at = idx(row);
    try {
      const xml = addRow(entry.xml ?? "", r, io);
      await guide.replaceTable({ start: at, guard: entry.text ?? "", xml, select: at });
    } catch (e) {
      ui.say(messageOf(e), true);
      return;
    }
    render();
    const fresh = rowAt(at);
    if (!fresh) return;
    if (toFirst) {
      const first = tableOf(fresh).rows[r + 1]?.cells.findIndex((cell) => cell.editable) ?? -1;
      if (first >= 0) startCellEdit(fresh, r + 1, first, "start");
      else focusCell(fresh, r + 1, 0);
    } else focusCell(fresh, r + 1, c);
    ui.say("Row added. Ctrl+Z takes it back.");
  }

  /** @param {Row} row @param {number} r @param {number} c */
  async function deleteTableRow(row, r, c) {
    await settle();
    stopCellEdit();
    const entry = entryOf(row);
    const at = idx(row);
    try {
      const xml = deleteRow(entry.xml ?? "", r, io);
      await guide.replaceTable({ start: at, guard: entry.text ?? "", xml, select: at });
    } catch (e) {
      ui.say(messageOf(e), true);
      return;
    }
    render();
    const fresh = rowAt(at);
    if (fresh) focusCell(fresh, r, c);
    ui.say("Row deleted. Ctrl+Z brings it back.");
  }

  /**
   * A change to the table's columns or cells, or to its title and description: made on the table's XML (src/guide-table.mjs) and put back as the table.
   * `c` is the cell the author is in, which is the place in its row; the column it is in is found from the cells before it.
   * @param {string} name @param {Row} row @param {number} r @param {number} c
   */
  async function tableChange(name, row, r, c) {
    await settle();
    stopCellEdit();
    const entry = entryOf(row);
    const at = idx(row);
    const xml0 = entry.xml ?? "";
    try {
      let xml = xml0;
      let said = "";
      // The grid column a cell begins at: the cells before it, each as wide as the columns it covers.
      const column = readTable(xml0, io).rows[r]?.cells.slice(0, c).reduce((n, cell) => n + (cell.span || 1), 0) ?? 0;
      if (name === "addColumn") {
        xml = addColumn(xml0, column + ((readTable(xml0, io).rows[r]?.cells[c]?.span || 1) - 1), io);
        said = "Column added after this one. Name it in the header row. Ctrl+Z takes it back.";
      } else if (name === "deleteColumn") {
        xml = deleteColumn(xml0, column, io);
        said = "Column taken out, with the words in it. Ctrl+Z brings it back.";
      } else if (name === "mergeCell") {
        xml = mergeCell(xml0, r, c, io);
        said = "This cell and the next are one cell now. Ctrl+Z takes it back.";
      } else if (name === "splitCell") {
        xml = splitCell(xml0, r, c, io);
        said = "The joined cell is cells again. Ctrl+Z takes it back.";
      } else {
        const now = readTable(xml0, io);
        if (!ui.ask) return ui.say("This cannot be changed here.", true);
        const answer = await ui.ask({
          title: "Table title and description",
          intro: "What a screen reader says about the table: its title first, then the description when the person asks for more.",
          confirm: "Change",
          fields: [
            { id: "caption", label: "Title (a few words naming the table)", type: "text", value: now.caption },
            { id: "description", label: "Description (what the table shows)", type: "lines", rows: 3, value: now.description, hint: `Up to ${ALT_LIMITS.description} letters. It may be left empty.` },
          ],
          finish: (/** @type {Record<string, any>} */ v) => {
            setTableAlt(xml0, { caption: String(v.caption), description: String(v.description) }, io);
            return { caption: String(v.caption), description: String(v.description) };
          },
        });
        if (!answer) {
          focusCell(row, r, c);
          return;
        }
        xml = setTableAlt(xml0, answer, io);
        said = "The table's title and description are changed. Ctrl+Z takes it back.";
      }
      await guide.replaceTable({ start: at, guard: entry.text ?? "", xml, select: at });
      render();
      const fresh = rowAt(at);
      const cells = fresh ? tableOf(fresh).rows[r]?.cells.length ?? 1 : 1;
      if (fresh) focusCell(fresh, r, Math.max(0, Math.min(c, cells - 1)));
      ui.say(said);
    } catch (e) {
      ui.say(messageOf(e), true);
    }
  }

  /** The picture in the paragraph the author is on: its description, whether it is only decoration, its width and the caption under it. */
  async function pictureDetails() {
    const rows = targetRows();
    if (rows.length !== 1) return ui.say("Choose a paragraph that holds a picture first.");
    const row = rows[0];
    await settle();
    const at = idx(row);
    const entry = entryOf(row);
    const now = readPicture(entry.xml ?? "", io);
    if (!now) return ui.say("This paragraph does not hold exactly one picture of yours. The icons of blocks are the program's own and are not changed.");
    if (!ui.ask) return ui.say("This cannot be changed here.", true);
    const below = entries[at + 1];
    const hasCaption = below?.style === "OSG Caption";
    const written = hasCaption ? String(below.text ?? "").replace(/^Figure\s*[\d.]*\s*:\s*/, "").trim() : "";
    stopEdit();
    const answer = await ui.ask({
      title: "Picture details",
      intro: "A description is what a screen reader says in place of the picture. A picture that is only decoration says nothing.",
      confirm: "Change",
      fields: [
        { id: "alt", label: "Describe the picture in a few words", type: "text", value: now.description, hint: `Up to ${DESCRIPTION_LIMIT} letters.` },
        { id: "decorative", label: "This picture is only decoration", type: "check", value: now.decorative },
        { id: "width", label: "Width", type: "select", value: now.width || PICTURE_WIDTHS[0].id, options: PICTURE_WIDTHS.map((w) => ({ value: w.id, label: w.label })) },
        { id: "caption", label: "Caption under the picture (optional)", type: "text", value: written, hint: "Update guide numbers captions: Figure 1.1, Figure 1.2." },
      ],
      finish: (/** @type {Record<string, any>} */ v) => {
        setPicture(entry.xml ?? "", { description: String(v.alt), decorative: v.decorative === true, width: String(v.width) }, io);
        if (String(v.caption).trim()) captionLine({ caption: String(v.caption) });
        return { alt: String(v.alt), decorative: v.decorative === true, width: String(v.width), caption: String(v.caption).trim() };
      },
    });
    if (!answer) {
      focusRow(row);
      return;
    }
    try {
      const xml = setPicture(entry.xml ?? "", { description: answer.alt, decorative: answer.decorative, width: answer.width }, io);
      await guide.replaceParagraphs({ start: at, count: 1, guard: entry.text ?? "", xmls: [xml], select: at });
      // The caption is its own paragraph in the caption style, straight under the picture: changed, added or taken away.
      const wanted = answer.caption;
      const nowBelow = (guide.exportState().paragraphs ?? [])[at + 1];
      const isCaption = nowBelow?.style === "OSG Caption";
      if (wanted) {
        const line = captionLine({ caption: wanted }).replace(/^<w:p>/, '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">');
        if (isCaption) await guide.replaceParagraphs({ start: at + 1, count: 1, guard: nowBelow.text ?? "", xmls: [line], select: at });
        else await guide.replaceParagraphs({ start: at + 1, count: 0, guard: "", xmls: [line], select: at });
      } else if (isCaption) await guide.replaceParagraphs({ start: at + 1, count: 1, guard: nowBelow.text ?? "", xmls: [], select: at });
    } catch (e) {
      ui.say(messageOf(e), true);
      return;
    }
    render();
    focusRow(rowAt(at));
    ui.say("The picture is changed. Ctrl+Z takes it back.");
  }

  /** Whether the keys belong to a box being typed in (a plain text box or the box for a paragraph's words). @param {EventTarget | null} target */
  function isTextBox(target) {
    return target instanceof HTMLTextAreaElement || (target instanceof HTMLElement && target.isContentEditable);
  }

  /** @param {EventTarget | null} target @returns {Row | null} */
  function rowOf(target) {
    const el = target instanceof Element ? target.closest("[data-index]") : null;
    return el ? shown.find((r) => r.el === el) ?? null : null;
  }

  // The keys of a row that has the keyboard, and a click on one.
  page.addEventListener("keydown", (ev) => {
    if (isTextBox(ev.target)) return;
    if (ev.target instanceof HTMLElement && ev.target.matches("td[data-row], th[data-row]")) {
      onCellKey(ev, ev.target);
      return;
    }
    const row = rowOf(ev.target);
    if (!row || ev.target !== row.el) return;
    // A key on a row (moving about, choosing) ends what was chosen in the words before.
    textSel = null;
    const action = rowKey({ key: ev.key, shift: ev.shiftKey, ctrl: ev.ctrlKey, alt: ev.altKey, meta: ev.metaKey });
    if (!action) return;
    ev.preventDefault();
    const at = shown.indexOf(row);
    if (action === "down") {
      clearSelection();
      focusRow(shown[at + 1] ?? row);
    } else if (action === "up") {
      clearSelection();
      focusRow(shown[at - 1] ?? row);
    } else if (action === "first") {
      clearSelection();
      focusRow(shown[0]);
    } else if (action === "last") {
      clearSelection();
      focusRow(shown[shown.length - 1]);
    } else if (action === "edit") startEdit(row, "end");
    else if (action === "extendDown" || action === "extendUp") extend(row, action === "extendDown" ? 1 : -1);
    else if (action === "clear") clearSelection();
    else if (action === "selectAll") {
      if (shown.length) select(idx(shown[0]), idx(shown[shown.length - 1]));
    } else if (action === "undo" || action === "redo") run(() => history(action === "redo", false));
    else run(() => command(action));
  });
  page.addEventListener("click", (ev) => {
    if (isTextBox(ev.target)) return;
    const cell = ev.target instanceof Element ? ev.target.closest("td[data-row], th[data-row]") : null;
    if (cell instanceof HTMLElement) {
      const inTable = rowOf(cell);
      if (inTable) {
        clearSelection();
        startCellEdit(inTable, Number(cell.dataset.row), Number(cell.dataset.col), "end");
      }
      return;
    }
    const row = rowOf(ev.target);
    if (!row) return;
    if (ev.shiftKey && lastRow && lastRow !== row && shown.includes(lastRow)) {
      select(selection ? selection.anchor : idx(lastRow), idx(row));
      focusRow(row);
      return;
    }
    clearSelection();
    startEdit(row, "end");
  });
  // Ctrl+V on a row: paragraphs cut or copied in the guide come back as they were, other text becomes a paragraph for each line.
  page.addEventListener("paste", (ev) => {
    if (isTextBox(ev.target)) return;
    ev.preventDefault();
    const text = ev.clipboardData?.getData("text/plain") ?? "";
    run(() => paste(text));
  });
  page.addEventListener("focusin", (ev) => {
    if (ev.target instanceof HTMLElement && ev.target.matches("td[data-row], th[data-row]")) {
      const inTable = rowOf(ev.target);
      if (inTable) {
        if (selection && !selectedRows().includes(inTable)) clearSelection();
        grid = { row: inTable, r: Number(ev.target.dataset.row), c: Number(ev.target.dataset.col) };
        lastRow = inTable;
        setTabStop(inTable);
        placeCursor(inTable);
      }
      return;
    }
    if (isTextBox(ev.target)) return;
    const row = rowOf(ev.target);
    if (row && ev.target === row.el) {
      // The keyboard going to a row that is not one of the chosen rows ends the choice (going along with the choice, as Shift+Down does, keeps it).
      if (selection && !selectedRows().includes(row)) clearSelection();
      setTabStop(row);
      lastRow = row;
      placeCursor(row);
    }
  });

  // ---- the toolbar: the looks of words and the alignments, as buttons (the Format menu has them too)
  /** @type {([string, string, string, string?] | "-")[]} name, words, icon, the setting a toggle shows */
  const TOOLS = [
    ["bold", "Bold", "bold", "bold"],
    ["italic", "Italic", "italic", "italic"],
    ["underline", "Underline", "underline", "underline"],
    ["strike", "Strikethrough", "strikethrough", "strike"],
    ["superscript", "Superscript", "superscript", "script:superscript"],
    ["subscript", "Subscript", "subscript", "script:subscript"],
    "-",
    ["alignLeft", "Align left", "text-align-start"],
    ["alignCenter", "Center", "text-align-center"],
    ["alignRight", "Align right", "text-align-end"],
    "-",
    ["text", "Text...", "type"],
    ["charStyle", "Character style...", "baseline"],
    ["clearText", "Clear text formatting", "remove-formatting"],
  ];
  /** @type {{ button: HTMLButtonElement, setting?: string, base?: string }[]} */
  const tools = [];
  for (const t of TOOLS) {
    if (t === "-") {
      const sep = doc.createElement("span");
      sep.className = "format-sep";
      sep.setAttribute("role", "separator");
      bar.append(sep);
      continue;
    }
    const [name, words, icon, setting] = t;
    const button = doc.createElement("button");
    button.type = "button";
    button.className = "format-tool";
    button.setAttribute("aria-label", words);
    button.title = words;
    button.append(uiIcon(doc, /** @type {any} */ (icon)));
    if (setting) button.setAttribute("aria-pressed", "false");
    button.tabIndex = tools.length === 0 ? 0 : -1;
    // The mouse does not take the keyboard from the words: the chosen words stay chosen.
    button.addEventListener("mousedown", (ev) => ev.preventDefault());
    button.addEventListener("click", () => void run(() => command(name)));
    bar.append(button);
    tools.push({ button, setting });
  }
  // Zoom of the guide's page, at the right end of the bar: smaller, the size it is (a button: back to 100%), bigger. The program's page owns it (src/zoom.mjs) and
  // says when it changes (an "osg-zoom" event on the page), so the number here is always the one on the page.
  let drawLevel = () => {};
  if (ui.zoom) {
    const zoom = ui.zoom;
    const group = doc.createElement("span");
    group.className = "format-zoom";
    group.setAttribute("role", "group");
    group.setAttribute("aria-label", "Zoom");
    /** @param {string} words @param {() => unknown} act @returns {HTMLButtonElement} */
    const zoomButton = (words, act) => {
      const button = doc.createElement("button");
      button.type = "button";
      button.className = "format-tool";
      button.setAttribute("aria-label", words);
      button.title = words;
      button.tabIndex = -1;
      button.addEventListener("mousedown", (ev) => ev.preventDefault());
      button.addEventListener("click", () => void act());
      return button;
    };
    const smaller = zoomButton("Zoom out", () => zoom.less());
    smaller.append(uiIcon(doc, "zoom-out"));
    const level = zoomButton("Actual size", () => zoom.actual());
    level.classList.add("format-zoom-level");
    const bigger = zoomButton("Zoom in", () => zoom.more());
    bigger.append(uiIcon(doc, "zoom-in"));
    group.append(smaller, level, bigger);
    bar.append(group);
    for (const button of [smaller, level, bigger]) tools.push({ button });
    drawLevel = () => {
      const percent = `${Math.round(zoom.level() * 100)}%`;
      level.textContent = percent;
      level.setAttribute("aria-label", `Zoom ${percent}. Choose to go back to the actual size, 100%.`);
      level.title = `Zoom ${percent}: choose to go back to 100%`;
    };
    drawLevel();
    doc.addEventListener("osg-zoom", drawZoomLevel);
  }
  function drawZoomLevel() {
    drawLevel();
  }
  // One Tab stop for the bar, and the arrow keys along it.
  bar.addEventListener("keydown", (ev) => {
    const at = tools.findIndex((t) => t.button === doc.activeElement);
    if (at < 0) return;
    const to = ev.key === "ArrowRight" ? at + 1 : ev.key === "ArrowLeft" ? at - 1 : ev.key === "Home" ? 0 : ev.key === "End" ? tools.length - 1 : -1;
    if (to < 0) return;
    ev.preventDefault();
    const next = tools[(to + tools.length) % tools.length];
    for (const t of tools) t.button.tabIndex = t === next ? 0 : -1;
    next.button.focus();
  });
  /** The buttons that switch a look show whether the chosen words have it. */
  function syncBar() {
    if (gone) return;
    const e = editing;
    /** @type {Record<string, any> | null} */
    let now = null;
    /** @type {Record<string, any>} */
    let base = {};
    if (e) {
      now = e.ta.state();
      base = baseFont(entryOf(e.row).style ?? "Normal");
    }
    for (const t of tools) {
      if (!t.setting) continue;
      let on = false;
      if (now) {
        if (t.setting.startsWith("script:")) on = now.script === t.setting.slice(7);
        else on = now[t.setting] === true || (now[t.setting] === undefined && base[t.setting] === true);
      }
      t.button.setAttribute("aria-pressed", String(on));
    }
  }
  doc.addEventListener("selectionchange", syncBar);

  render();
  return {
    render,
    /** The document is about to go: stop everything. */
    destroy() {
      gone = true;
      if (selectionTimer !== null) window.clearTimeout(selectionTimer);
      doc.removeEventListener("selectionchange", syncBar);
      doc.removeEventListener("osg-zoom", drawZoomLevel);
      editing = null;
    },
    /** A command of the Edit menu, for the rows the author is on or has chosen. @param {string} name */
    command(name) {
      return run(() => (name === "undo" || name === "redo" ? history(name === "redo", false) : command(name)));
    },
    /** Put a block after the row the author is on. @param {(ctx: any) => string} build @param {string} what */
    insert(build, what) {
      return run(() => insertBlock(build, what));
    },
    /** What the Edit menu needs to know to dim what cannot be done now. */
    state() {
      const h = guide.historyState();
      return { canUndo: h.undo > 0, canRedo: h.redo > 0, hasRows: targetRows().length > 0, inTable: !!grid && shown.includes(grid.row), onPicture: onPicture() };
    },
    /** Put the keyboard on the row of the cursor. */
    focusCursor() {
      focusRow(rowNear(guide.cursor()));
    },
  };
}
