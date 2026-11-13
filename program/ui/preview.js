// @ts-check
/**
 * The guide as it reads, drawn from the paragraphs of the project (their Word-format XML and styles). A view, not yet an editor. It is made in
 * two steps so it can be tested without a page: `describeGuide` turns the paragraphs into a tree of plain objects, and `renderGuide` draws the
 * tree with the page's own document, using text and attributes only (never markup from the guide).
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const REF = "osg-ref:";

/**
 * @typedef {{ tag: string, cls?: string, text?: string, attrs?: Record<string, string>, height?: number, index?: number, children?: Node_[] }} Node_
 * @typedef {{ xml?: string, text?: string, style?: string, builtIn?: string }} Entry
 */

/** @param {Element} el @param {string} name */
const kids = (el, name) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && /** @type {Element} */ (n).localName === name);
/** @param {Element} el @param {string} ns @param {string} name */
const first = (el, ns, name) => Array.from(el.getElementsByTagNameNS(ns, name))[0];

/**
 * The inline parts of a paragraph, in order: {t:'text',text,label}, {t:'tab'}, {t:'br'}, {t:'check'}, {t:'cite',text}.
 * @param {Element} el
 * @returns {{ t: string, text?: string, label?: boolean }[]}
 */
function inlineParts(el) {
  /** @type {{ t: string, text?: string, label?: boolean }[]} */
  const out = [];
  const walk = (/** @type {Element} */ node) => {
    for (const c of Array.from(node.childNodes)) {
      if (c.nodeType !== 1) continue;
      const e = /** @type {Element} */ (c);
      if (e.namespaceURI !== W && e.namespaceURI !== W14) continue;
      const name = e.localName;
      if (name === "r") {
        const rstyle = first(e, W, "rStyle")?.getAttributeNS(W, "val");
        for (const rc of Array.from(e.childNodes)) {
          if (rc.nodeType !== 1) continue;
          const r = /** @type {Element} */ (rc);
          if (r.localName === "t") out.push({ t: "text", text: r.textContent ?? "", label: rstyle === "OSGCueLabel" });
          else if (r.localName === "tab") out.push({ t: "tab" });
          else if (r.localName === "br") out.push({ t: "br" });
        }
      } else if (name === "sdt") {
        const pr = kids(e, "sdtPr")[0];
        const tag = pr ? first(pr, W, "tag")?.getAttributeNS(W, "val") ?? "" : "";
        const content = kids(e, "sdtContent")[0];
        if (pr && first(pr, W14, "checkbox")) out.push({ t: "check" });
        else if (tag.startsWith(REF) && content) out.push({ t: "cite", text: textOfContent(content) });
        else if (content) walk(content);
      } else if (["hyperlink", "smartTag", "ins", "fldSimple", "customXml"].includes(name)) walk(e);
    }
  };
  walk(el);
  return out;
}

/** @param {Element} el */
function textOfContent(el) {
  return Array.from(el.getElementsByTagNameNS(W, "t")).map((t) => t.textContent ?? "").join("");
}

/** @param {{ t: string, text?: string, label?: boolean }[]} parts @returns {Node_[]} */
function inlineNodes(parts) {
  /** @type {Node_[]} */
  const out = [];
  for (const p of parts) {
    if (p.t === "text" && p.text) out.push({ tag: "span", cls: p.label ? "g-label" : undefined, text: p.text });
    else if (p.t === "cite") out.push({ tag: "span", cls: "g-cite", text: p.text });
    else if (p.t === "check") out.push({ tag: "span", cls: "g-check", text: "☐" });
    else if (p.t === "br") out.push({ tag: "br" });
    else if (p.t === "tab") out.push({ tag: "span", text: " " });
  }
  return out;
}

/** A field row: each label followed by a line to write on. @param {{ t: string, text?: string }[]} parts @returns {Node_[]} */
function fieldNodes(parts) {
  /** @type {Node_[]} */
  const out = [];
  let label = "";
  const flush = () => {
    if (label.trim()) out.push({ tag: "span", cls: "g-fieldlabel", text: label.trim() }, { tag: "span", cls: "g-fieldline" });
    label = "";
  };
  for (const p of parts) {
    if (p.t === "text") label += p.text ?? "";
    else if (p.t === "tab") flush();
  }
  flush();
  return out;
}

/** @param {Element} p @returns {number} height in points of an exact-height paragraph, or 0 */
function exactHeight(p) {
  const s = first(p, W, "spacing");
  const line = Number(s?.getAttributeNS(W, "line") ?? 0);
  return s?.getAttributeNS(W, "lineRule") === "exact" && line > 0 ? line / 20 : 0;
}

/** @type {Record<string, { tag: string, cls: string }>} */
const BY_STYLE = {
  "OSG Guide Title": { tag: "h1", cls: "g-title" },
  "OSG Module": { tag: "h2", cls: "g-module" },
  "OSG Assessment": { tag: "h2", cls: "g-module" },
  "OSG Lesson": { tag: "h3", cls: "g-lesson" },
  "OSG Time": { tag: "p", cls: "g-time" },
  "OSG Cue": { tag: "div", cls: "g-cue" },
  "OSG Cue Facilitator": { tag: "div", cls: "g-cue facilitator" },
  "OSG Cue Body": { tag: "p", cls: "g-body" },
  "OSG Facilitator Notes": { tag: "div", cls: "g-notes" },
  "OSG Answer Key": { tag: "div", cls: "g-notes" },
  "OSG Form Title": { tag: "h4", cls: "g-form" },
  "OSG Form Title Facilitator": { tag: "h4", cls: "g-form facilitator" },
  "OSG Form Text": { tag: "p", cls: "g-formtext" },
  "OSG Form Source": { tag: "p", cls: "g-source" },
  "OSG Form Print": { tag: "p", cls: "g-print" },
  "OSG Form Card": { tag: "div", cls: "g-card" },
  "heading 1": { tag: "h2", cls: "g-heading" },
  "heading 2": { tag: "h3", cls: "g-heading2" },
  Title: { tag: "h1", cls: "g-title" },
};

/** A paragraph with nothing in it: no words, no tab, no check box, no picture. @param {Element} p */
function isBlank(p) {
  return inlineParts(p).length === 0 && p.getElementsByTagNameNS(W, "drawing").length === 0;
}

/** The row for a paragraph with nothing in it, in the look of its style, with words for a screen reader. @param {string} style @returns {Node_} */
function blankNode(style) {
  const known = BY_STYLE[style];
  return { tag: known?.tag ?? "p", cls: `${known?.cls ?? "g-para"} g-blank`, children: [{ tag: "span", cls: "sr-only", text: "Empty paragraph" }] };
}

/**
 * The tree for one paragraph (a w:p element and its style's name).
 * @param {Element} p @param {string} style @returns {Node_ | null}
 */
function paragraphNode(p, style) {
  const parts = inlineParts(p);
  const words = parts.filter((x) => x.t === "text").map((x) => x.text).join("").trim();
  if (style === "OSG Form Line" || style === "OSG Answer Line") return { tag: "div", cls: "g-rule" };
  if (style === "OSG Form Box") return { tag: "div", cls: "g-box", height: exactHeight(p) || 60 };
  if (style === "OSG Form Field") return { tag: "div", cls: "g-field", children: fieldNodes(parts) };
  if (!words && !parts.some((x) => x.t === "check" || x.t === "cite")) return null;
  const known = BY_STYLE[style];
  if (known) return { tag: known.tag, cls: known.cls, children: inlineNodes(parts) };
  if (/^toc \d$/i.test(style)) return { tag: "p", cls: "g-toc", children: inlineNodes(parts) };
  return { tag: "p", cls: /^OSG /.test(style) ? "g-para" : "g-para", children: inlineNodes(parts) };
}

/**
 * A table: a header row where Word's header-row mark says so, cells with the words of their paragraphs.
 * @param {Element} tbl @returns {Node_}
 */
function tableNode(tbl) {
  const caption = first(tbl, W, "tblCaption")?.getAttributeNS(W, "val") ?? "";
  /** @type {Node_[]} */
  const rows = [];
  for (const tr of kids(tbl, "tr")) {
    const header = !!first(tr, W, "tblHeader") && kids(tr, "trPr").length > 0;
    /** @type {Node_[]} */
    const cells = [];
    for (const tc of kids(tr, "tc")) {
      /** @type {Node_[]} */
      const content = [];
      for (const para of kids(tc, "p")) {
        const parts = inlineParts(para);
        if (!parts.length) continue;
        if (content.length) content.push({ tag: "br" });
        content.push(...inlineNodes(parts));
      }
      cells.push({ tag: header ? "th" : "td", attrs: header ? { scope: "col" } : undefined, children: content });
    }
    rows.push({ tag: "tr", children: cells });
  }
  return { tag: "table", cls: "g-table", attrs: caption ? { "aria-label": caption } : undefined, children: rows };
}

/**
 * The guide's paragraphs as a tree of plain objects, each top-level node carrying the index of the paragraph it is drawn from. A table's XML
 * rides on its first cell's entry and the others carry none, so an entry with no XML that follows a table is part of it.
 * With `blank`, a paragraph with nothing in it is drawn as an empty row (the editor needs somewhere to type); without it, it is left out.
 * @param {Entry[]} entries
 * @param {(xml: string) => Document} parseXml
 * @param {{ blank?: boolean }} [options]
 * @returns {Node_[]}
 */
export function describeGuide(entries, parseXml, options = {}) {
  /** @type {Node_[]} */
  const out = [];
  for (let index = 0; index < entries.length; index++) {
    const entry = entries[index];
    const xml = entry.xml ?? "";
    if (!xml.trim()) continue;
    let root = null;
    try {
      root = parseXml(xml).documentElement;
    } catch {
      root = null;
    }
    if (!root) {
      out.push({ tag: "p", cls: "g-other", text: entry.text ?? "", index });
      continue;
    }
    if (root.localName === "tbl") out.push({ ...tableNode(root), index });
    else if (root.localName === "p") {
      const node = paragraphNode(root, entry.style ?? "Normal");
      if (node) out.push({ ...node, index });
      else if (options.blank && isBlank(root)) out.push({ ...blankNode(entry.style ?? "Normal"), index });
    }
  }
  return out;
}

/**
 * Draw a tree into `parent`, with the page's own document: text and attributes only.
 * @param {Document} doc @param {Element} parent @param {Node_[]} nodes
 */
export function drawNodes(doc, parent, nodes) {
  for (const n of nodes) {
    const el = doc.createElement(n.tag);
    if (n.cls) el.className = n.cls;
    for (const [k, v] of Object.entries(n.attrs ?? {})) el.setAttribute(k, v);
    if (n.height) /** @type {HTMLElement} */ (el).style.height = `${Math.round(n.height * 1.33)}px`;
    if (n.text !== undefined) el.textContent = n.text;
    if (n.children) drawNodes(doc, el, n.children);
    parent.append(el);
  }
}

/**
 * Draw the guide into `container`, replacing what was there.
 * @param {Document} doc @param {Element} container @param {Entry[]} entries @param {(xml: string) => Document} parseXml
 */
export function renderGuide(doc, container, entries, parseXml) {
  container.replaceChildren();
  const nodes = describeGuide(entries, parseXml);
  const page = doc.createElement("article");
  page.className = "page";
  page.setAttribute("aria-label", "The guide");
  if (!nodes.length) {
    const p = doc.createElement("p");
    p.className = "g-empty";
    p.textContent = "This guide is empty. Choose Guide overview in the tools beside it to start it.";
    page.append(p);
  } else drawNodes(doc, page, nodes);
  container.append(page);
}
