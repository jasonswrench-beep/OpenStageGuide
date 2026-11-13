// @ts-check
// SPDX-License-Identifier: Apache-2.0 AND ISC
// Made by scripts/make-ui-icons.mjs from program/icons.json and the Lucide shapes OpenStage Author ships (editor/web/icons/lucide.json):
// do not edit by hand. Change the list and run the script again.
// The icons of the program's own buttons, menus and command search: Lucide 1.56.0, kept as their drawn shapes, drawn as lines in the
// colour of the words around them, and hidden from screen readers (the words beside an icon, or a button's name, say what it is).
//
// ISC License
//
// Copyright (c) 2026 Lucide Icons and Contributors
//
// Permission to use, copy, modify, and/or distribute this software for any
// purpose with or without fee is hereby granted, provided that the above
// copyright notice and this permission notice appear in all copies.
//
// THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
// WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
// MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
// ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
// WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
// ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
// OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
//
// ---
//
// The following Lucide icons are derived from the Feather project:
//
// airplay, alert-circle, alert-octagon, alert-triangle, aperture, arrow-down-circle, arrow-down-left, arrow-down-right, arrow-down, arrow-left-circle, arrow-left, arrow-right-circle, arrow-right, arrow-up-circle, arrow-up-left, arrow-up-right, arrow-up, at-sign, calendar, cast, check, chevron-down, chevron-left, chevron-right, chevron-up, chevrons-down, chevrons-left, chevrons-right, chevrons-up, circle, clipboard, clock, code, columns, command, compass, corner-down-left, corner-down-right, corner-left-down, corner-left-up, corner-right-down, corner-right-up, corner-up-left, corner-up-right, crosshair, database, divide-circle, divide-square, dollar-sign, download, external-link, feather, frown, hash, headphones, help-circle, info, italic, key, layout, life-buoy, link-2, link, loader, lock, log-in, log-out, maximize, meh, minimize, minimize-2, minus-circle, minus-square, minus, monitor, moon, more-horizontal, more-vertical, move, music, navigation-2, navigation, octagon, pause-circle, percent, plus-circle, plus-square, plus, power, radio, rss, search, server, share, shopping-bag, sidebar, smartphone, smile, square, table-2, tablet, target, terminal, trash-2, trash, triangle, tv, type, upload, x-circle, x-octagon, x-square, x, zoom-in, zoom-out
//
// The MIT License (MIT) (for the icons listed above)
//
// Copyright (c) 2013-present Cole Bemis
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

/** @typedef {{ tag: string } & Record<string, string | number>} Shape */

export const UI_ICONS = /** @type {const} */ ({
  "archive": [{"tag":"rect","width":20,"height":5,"x":2,"y":3,"rx":1},{"tag":"path","d":"M4 8v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8"},{"tag":"path","d":"M10 12h4"}],
  "archive-restore": [{"tag":"rect","width":20,"height":5,"x":2,"y":3,"rx":1},{"tag":"path","d":"M4 8v11a2 2 0 0 0 2 2h2"},{"tag":"path","d":"M20 8v11a2 2 0 0 1-2 2h-2"},{"tag":"path","d":"m9 15 3-3 3 3"},{"tag":"path","d":"M12 12v9"}],
  "arrow-down": [{"tag":"path","d":"M12 5v14"},{"tag":"path","d":"m19 12-7 7-7-7"}],
  "arrow-up": [{"tag":"path","d":"m5 12 7-7 7 7"},{"tag":"path","d":"M12 19V5"}],
  "baseline": [{"tag":"path","d":"M4 20h16"},{"tag":"path","d":"m6 16 6-12 6 12"},{"tag":"path","d":"M8 12h8"}],
  "bold": [{"tag":"path","d":"M6 12h9a4 4 0 0 1 0 8H7a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h7a4 4 0 0 1 0 8"}],
  "book-open": [{"tag":"path","d":"M12 5v16"},{"tag":"path","d":"M20.001 19A2 2 0 0022 17V5a2 2 0 00-1.999-2L16 3.002A5 5 0 0012 5a5 5 0 00-4-2H4a2 2 0 00-2 2v12a2 2 0 001.999 2H8a5 5 0 014 2 5 5 0 014-2z"}],
  "book-open-check": [{"tag":"path","d":"M12 5v16"},{"tag":"path","d":"m16 12 2 2 4-4"},{"tag":"path","d":"M22 6V5a2 2 0 00-1.999-2L16 3.002A5 5 0 0012 5a5 5 0 00-4-2H4a2 2 0 00-2 2v12a2 2 0 001.999 2H8a5 5 0 014 2 5 5 0 014-2h4.001A2 2 0 0022 17v-1.344"}],
  "check": [{"tag":"path","d":"M20 6 9 17l-5-5"}],
  "chevron-down": [{"tag":"path","d":"m6 9 6 6 6-6"}],
  "circle-alert": [{"tag":"circle","cx":12,"cy":12,"r":10},{"tag":"line","x1":12,"x2":12,"y1":8,"y2":12},{"tag":"line","x1":12,"x2":12.01,"y1":16,"y2":16}],
  "circle-check": [{"tag":"circle","cx":12,"cy":12,"r":10},{"tag":"path","d":"m16 9-5.5 5.5L8 12"}],
  "circle-question-mark": [{"tag":"circle","cx":12,"cy":12,"r":10},{"tag":"path","d":"M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"},{"tag":"path","d":"M12 17h.01"}],
  "clipboard-list": [{"tag":"rect","width":8,"height":4,"x":8,"y":2,"rx":1,"ry":1},{"tag":"path","d":"M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"},{"tag":"path","d":"M12 11h4"},{"tag":"path","d":"M12 16h4"},{"tag":"path","d":"M8 11h.01"},{"tag":"path","d":"M8 16h.01"}],
  "clipboard-paste": [{"tag":"path","d":"M11 14h10"},{"tag":"path","d":"M16 4h2a2 2 0 0 1 2 2v1.344"},{"tag":"path","d":"m17 18 4-4-4-4"},{"tag":"path","d":"M8 4H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 1.793-1.113"},{"tag":"rect","x":8,"y":2,"width":8,"height":4,"rx":1}],
  "clock": [{"tag":"circle","cx":12,"cy":12,"r":10},{"tag":"path","d":"M12 6v6l4 2"}],
  "copy": [{"tag":"rect","width":14,"height":14,"x":8,"y":8,"rx":2,"ry":2},{"tag":"path","d":"M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"}],
  "ellipsis": [{"tag":"circle","cx":12,"cy":12,"r":1},{"tag":"circle","cx":19,"cy":12,"r":1},{"tag":"circle","cx":5,"cy":12,"r":1}],
  "eraser": [{"tag":"path","d":"M21 21H8a2 2 0 0 1-1.42-.587l-3.994-3.999a2 2 0 0 1 0-2.828l10-10a2 2 0 0 1 2.829 0l5.999 6a2 2 0 0 1 0 2.828L12.834 21"},{"tag":"path","d":"m5.082 11.09 8.828 8.828"}],
  "file-check": [{"tag":"path","d":"M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"},{"tag":"path","d":"M14 2v5a1 1 0 0 0 1 1h5"},{"tag":"path","d":"m9 15 2 2 4-4"}],
  "file-down": [{"tag":"path","d":"M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"},{"tag":"path","d":"M14 2v5a1 1 0 0 0 1 1h5"},{"tag":"path","d":"M12 18v-6"},{"tag":"path","d":"m9 15 3 3 3-3"}],
  "file-input": [{"tag":"path","d":"M4 11V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.706.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-1"},{"tag":"path","d":"M14 2v5a1 1 0 0 0 1 1h5"},{"tag":"path","d":"M2 15h10"},{"tag":"path","d":"m9 18 3-3-3-3"}],
  "file-plus": [{"tag":"path","d":"M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"},{"tag":"path","d":"M14 2v5a1 1 0 0 0 1 1h5"},{"tag":"path","d":"M9 15h6"},{"tag":"path","d":"M12 18v-6"}],
  "file-stack": [{"tag":"path","d":"M11 21a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1v-8a1 1 0 0 1 1-1"},{"tag":"path","d":"M16 16a1 1 0 0 1-1 1H9a1 1 0 0 1-1-1V8a1 1 0 0 1 1-1"},{"tag":"path","d":"M21 6a2 2 0 0 0-.586-1.414l-2-2A2 2 0 0 0 17 2h-3a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1z"}],
  "file-text": [{"tag":"path","d":"M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"},{"tag":"path","d":"M14 2v5a1 1 0 0 0 1 1h5"},{"tag":"path","d":"M10 9H8"},{"tag":"path","d":"M16 13H8"},{"tag":"path","d":"M16 17H8"}],
  "folder-open": [{"tag":"path","d":"m6 14 1.5-2.9A2 2 0 0 1 9.24 10H20a2 2 0 0 1 1.94 2.5l-1.54 6a2 2 0 0 1-1.95 1.5H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3.9a2 2 0 0 1 1.69.9l.81 1.2a2 2 0 0 0 1.67.9H18a2 2 0 0 1 2 2v2"}],
  "folder-plus": [{"tag":"path","d":"M12 10v6"},{"tag":"path","d":"M9 13h6"},{"tag":"path","d":"M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"}],
  "folders": [{"tag":"path","d":"M20 5a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h2.5a1.5 1.5 0 0 1 1.2.6l.6.8a1.5 1.5 0 0 0 1.2.6z"},{"tag":"path","d":"M3 8.268a2 2 0 0 0-1 1.738V19a2 2 0 0 0 2 2h11a2 2 0 0 0 1.732-1"}],
  "grid-3x3": [{"tag":"rect","width":18,"height":18,"x":3,"y":3,"rx":2},{"tag":"path","d":"M3 9h18"},{"tag":"path","d":"M3 15h18"},{"tag":"path","d":"M9 3v18"},{"tag":"path","d":"M15 3v18"}],
  "house": [{"tag":"path","d":"M15 21v-8a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v8"},{"tag":"path","d":"M3 10a2 2 0 0 1 .709-1.528l7-6a2 2 0 0 1 2.582 0l7 6A2 2 0 0 1 21 10v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"}],
  "id-card": [{"tag":"path","d":"M13 19a4 4 0 00-8 0"},{"tag":"path","d":"M16 10h2"},{"tag":"path","d":"M16 14h2"},{"tag":"circle","cx":9,"cy":12,"r":3},{"tag":"rect","x":2,"y":5,"width":20,"height":14,"rx":2}],
  "image": [{"tag":"rect","width":18,"height":18,"x":3,"y":3,"rx":2,"ry":2},{"tag":"circle","cx":9,"cy":9,"r":2},{"tag":"path","d":"m21 15-3.086-3.086a2 2 0 0 0-2.828 0L6 21"}],
  "info": [{"tag":"circle","cx":12,"cy":12,"r":10},{"tag":"path","d":"M12 16v-4"},{"tag":"path","d":"M12 8h.01"}],
  "italic": [{"tag":"line","x1":19,"x2":10,"y1":4,"y2":4},{"tag":"line","x1":14,"x2":5,"y1":20,"y2":20},{"tag":"line","x1":15,"x2":9,"y1":4,"y2":20}],
  "keyboard": [{"tag":"path","d":"M10 8h.01"},{"tag":"path","d":"M12 12h.01"},{"tag":"path","d":"M14 8h.01"},{"tag":"path","d":"M16 12h.01"},{"tag":"path","d":"M18 8h.01"},{"tag":"path","d":"M6 8h.01"},{"tag":"path","d":"M7 16h10"},{"tag":"path","d":"M8 12h.01"},{"tag":"rect","width":20,"height":16,"x":2,"y":4,"rx":2}],
  "layers": [{"tag":"path","d":"M12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83z"},{"tag":"path","d":"M2 12a1 1 0 0 0 .58.91l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 12"},{"tag":"path","d":"M2 17a1 1 0 0 0 .58.91l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 17"}],
  "layout-dashboard": [{"tag":"rect","width":7,"height":9,"x":3,"y":3,"rx":1},{"tag":"rect","width":7,"height":5,"x":14,"y":3,"rx":1},{"tag":"rect","width":7,"height":9,"x":14,"y":12,"rx":1},{"tag":"rect","width":7,"height":5,"x":3,"y":16,"rx":1}],
  "library": [{"tag":"path","d":"m16 6 4 14"},{"tag":"path","d":"M12 6v14"},{"tag":"path","d":"M8 8v12"},{"tag":"path","d":"M4 4v16"}],
  "list-checks": [{"tag":"path","d":"M13 5h8"},{"tag":"path","d":"M13 12h8"},{"tag":"path","d":"M13 19h8"},{"tag":"path","d":"m3 17 2 2 4-4"},{"tag":"path","d":"m3 7 2 2 4-4"}],
  "list-todo": [{"tag":"path","d":"M13 5h8"},{"tag":"path","d":"M13 12h8"},{"tag":"path","d":"M13 19h8"},{"tag":"path","d":"m3 17 2 2 4-4"},{"tag":"rect","x":3,"y":4,"width":6,"height":6,"rx":1}],
  "list-tree": [{"tag":"path","d":"M8 5h13"},{"tag":"path","d":"M13 12h8"},{"tag":"path","d":"M13 19h8"},{"tag":"path","d":"M3 10a2 2 0 0 0 2 2h3"},{"tag":"path","d":"M3 5v12a2 2 0 0 0 2 2h3"}],
  "message-square": [{"tag":"path","d":"M22 17a2 2 0 0 1-2 2H6.828a2 2 0 0 0-1.414.586l-2.202 2.202A.71.71 0 0 1 2 21.286V5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2z"}],
  "monitor-play": [{"tag":"path","d":"M15.033 9.44a.647.647 0 0 1 0 1.12l-4.065 2.352a.645.645 0 0 1-.968-.56V7.648a.645.645 0 0 1 .967-.56z"},{"tag":"path","d":"M12 17v4"},{"tag":"path","d":"M8 21h8"},{"tag":"rect","x":2,"y":3,"width":20,"height":14,"rx":2}],
  "palette": [{"tag":"path","d":"M12 22a1 1 0 0 1 0-20 10 9 0 0 1 10 9 5 5 0 0 1-5 5h-2.25a1.75 1.75 0 0 0-1.4 2.8l.3.4a1.75 1.75 0 0 1-1.4 2.8z"},{"tag":"circle","cx":13.5,"cy":6.5,"r":0.5,"fill":"currentColor"},{"tag":"circle","cx":17.5,"cy":10.5,"r":0.5,"fill":"currentColor"},{"tag":"circle","cx":6.5,"cy":12.5,"r":0.5,"fill":"currentColor"},{"tag":"circle","cx":8.5,"cy":7.5,"r":0.5,"fill":"currentColor"}],
  "pencil": [{"tag":"path","d":"M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z"},{"tag":"path","d":"m15 5 4 4"}],
  "pencil-line": [{"tag":"path","d":"M13 21h8"},{"tag":"path","d":"m15 5 4 4"},{"tag":"path","d":"M21.174 6.812a1 1 0 0 0-3.986-3.987L3.842 16.174a2 2 0 0 0-.5.83l-1.321 4.352a.5.5 0 0 0 .623.622l4.353-1.32a2 2 0 0 0 .83-.497z"}],
  "pilcrow": [{"tag":"path","d":"M13 4v16"},{"tag":"path","d":"M17 4v16"},{"tag":"path","d":"M19 4H9.5a4.5 4.5 0 0 0 0 9H13"}],
  "plus": [{"tag":"path","d":"M5 12h14"},{"tag":"path","d":"M12 5v14"}],
  "presentation": [{"tag":"path","d":"M2 3h20"},{"tag":"path","d":"M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3"},{"tag":"path","d":"m7 21 5-5 5 5"}],
  "printer": [{"tag":"path","d":"M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"},{"tag":"path","d":"M6 9V3a1 1 0 0 1 1-1h10a1 1 0 0 1 1 1v6"},{"tag":"rect","x":6,"y":14,"width":12,"height":8,"rx":1}],
  "quote": [{"tag":"path","d":"M16 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"},{"tag":"path","d":"M5 3a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2 1 1 0 0 1 1 1v1a2 2 0 0 1-2 2 1 1 0 0 0-1 1v2a1 1 0 0 0 1 1 6 6 0 0 0 6-6V5a2 2 0 0 0-2-2z"}],
  "redo-2": [{"tag":"path","d":"m15 14 5-5-5-5"},{"tag":"path","d":"M20 9H9.5A5.5 5.5 0 0 0 4 14.5A5.5 5.5 0 0 0 9.5 20H13"}],
  "refresh-cw": [{"tag":"path","d":"M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"},{"tag":"path","d":"M21 3v5h-5"},{"tag":"path","d":"M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"},{"tag":"path","d":"M8 16H3v5"}],
  "remove-formatting": [{"tag":"path","d":"M4 7V4h16v3"},{"tag":"path","d":"M5 20h6"},{"tag":"path","d":"M13 4 8 20"},{"tag":"path","d":"m15 15 5 5"},{"tag":"path","d":"m20 15-5 5"}],
  "scissors": [{"tag":"circle","cx":6,"cy":6,"r":3},{"tag":"path","d":"M8.12 8.12 12 12"},{"tag":"path","d":"M20 4 8.12 15.88"},{"tag":"circle","cx":6,"cy":18,"r":3},{"tag":"path","d":"M14.8 14.8 20 20"}],
  "search": [{"tag":"path","d":"m21 21-4.34-4.34"},{"tag":"circle","cx":11,"cy":11,"r":8}],
  "settings": [{"tag":"path","d":"M9.671 4.136a2.34 2.34 0 0 1 4.659 0 2.34 2.34 0 0 0 3.319 1.915 2.34 2.34 0 0 1 2.33 4.033 2.34 2.34 0 0 0 0 3.831 2.34 2.34 0 0 1-2.33 4.033 2.34 2.34 0 0 0-3.319 1.915 2.34 2.34 0 0 1-4.659 0 2.34 2.34 0 0 0-3.32-1.915 2.34 2.34 0 0 1-2.33-4.033 2.34 2.34 0 0 0 0-3.831A2.34 2.34 0 0 1 6.35 6.051a2.34 2.34 0 0 0 3.319-1.915"},{"tag":"circle","cx":12,"cy":12,"r":3}],
  "square": [{"tag":"rect","width":18,"height":18,"x":3,"y":3,"rx":2}],
  "square-pen": [{"tag":"path","d":"M12 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"},{"tag":"path","d":"M18.375 2.625a1 1 0 0 1 3 3l-9.013 9.014a2 2 0 0 1-.853.505l-2.873.84a.5.5 0 0 1-.62-.62l.84-2.873a2 2 0 0 1 .506-.852z"}],
  "square-plus": [{"tag":"rect","width":18,"height":18,"x":3,"y":3,"rx":2},{"tag":"path","d":"M8 12h8"},{"tag":"path","d":"M12 8v8"}],
  "strikethrough": [{"tag":"path","d":"M16 4H9a3 3 0 0 0-2.83 4"},{"tag":"path","d":"M14 12a4 4 0 0 1 0 8H6"},{"tag":"line","x1":4,"x2":20,"y1":12,"y2":12}],
  "subscript": [{"tag":"path","d":"m4 5 8 8"},{"tag":"path","d":"m12 5-8 8"},{"tag":"path","d":"M20 19h-4c0-1.5.44-2 1.5-2.5S20 15.33 20 14c0-.47-.17-.93-.48-1.29a2.11 2.11 0 0 0-2.62-.44c-.42.24-.74.62-.9 1.07"}],
  "superscript": [{"tag":"path","d":"m4 19 8-8"},{"tag":"path","d":"m12 19-8-8"},{"tag":"path","d":"M20 12h-4c0-1.5.442-2 1.5-2.5S20 8.334 20 7.002c0-.472-.17-.93-.484-1.29a2.105 2.105 0 0 0-2.617-.436c-.42.239-.738.614-.899 1.06"}],
  "swatch-book": [{"tag":"path","d":"M11 17a4 4 0 0 1-8 0V5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2Z"},{"tag":"path","d":"M16.7 13H19a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2H7"},{"tag":"path","d":"M 7 17h.01"},{"tag":"path","d":"m11 8 2.3-2.3a2.4 2.4 0 0 1 3.404.004L18.6 7.6a2.4 2.4 0 0 1 .026 3.434L9.9 19.8"}],
  "table": [{"tag":"path","d":"M12 3v18"},{"tag":"rect","width":18,"height":18,"x":3,"y":3,"rx":2},{"tag":"path","d":"M3 9h18"},{"tag":"path","d":"M3 15h18"}],
  "target": [{"tag":"circle","cx":12,"cy":12,"r":10},{"tag":"circle","cx":12,"cy":12,"r":6},{"tag":"circle","cx":12,"cy":12,"r":2}],
  "text-align-center": [{"tag":"path","d":"M21 5H3"},{"tag":"path","d":"M17 12H7"},{"tag":"path","d":"M19 19H5"}],
  "text-align-end": [{"tag":"path","d":"M21 5H3"},{"tag":"path","d":"M21 12H9"},{"tag":"path","d":"M21 19H7"}],
  "text-align-justify": [{"tag":"path","d":"M3 5h18"},{"tag":"path","d":"M3 12h18"},{"tag":"path","d":"M3 19h18"}],
  "text-align-start": [{"tag":"path","d":"M21 5H3"},{"tag":"path","d":"M15 12H3"},{"tag":"path","d":"M17 19H3"}],
  "text-cursor-input": [{"tag":"path","d":"M12 20h-1a2 2 0 0 1-2-2 2 2 0 0 1-2 2H6"},{"tag":"path","d":"M13 8h7a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-7"},{"tag":"path","d":"M5 16H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h1"},{"tag":"path","d":"M6 4h1a2 2 0 0 1 2 2 2 2 0 0 1 2-2h1"},{"tag":"path","d":"M9 6v12"}],
  "trash": [{"tag":"path","d":"M10 11v6"},{"tag":"path","d":"M14 11v6"},{"tag":"path","d":"M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"},{"tag":"path","d":"M3 6h18"},{"tag":"path","d":"M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"}],
  "triangle-alert": [{"tag":"path","d":"m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"},{"tag":"path","d":"M12 9v4"},{"tag":"path","d":"M12 17h.01"}],
  "type": [{"tag":"path","d":"M12 4v16"},{"tag":"path","d":"M4 7V5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v2"},{"tag":"path","d":"M9 20h6"}],
  "underline": [{"tag":"path","d":"M6 4v6a6 6 0 0 0 12 0V4"},{"tag":"line","x1":4,"x2":20,"y1":20,"y2":20}],
  "undo-2": [{"tag":"path","d":"M9 14 4 9l5-5"},{"tag":"path","d":"M4 9h10.5a5.5 5.5 0 0 1 5.5 5.5a5.5 5.5 0 0 1-5.5 5.5H11"}],
  "x": [{"tag":"path","d":"M18 6 6 18"},{"tag":"path","d":"m6 6 12 12"}],
});

/** @typedef {keyof typeof UI_ICONS} UiIconName */

// The name of the kind of element, which the browser needs to make one; it is never fetched.
const SVG_NS = "http://www.w3.org/2000/svg";

/**
 * An icon, drawn in the colour of the words beside it and hidden from screen readers.
 * @param {Document} doc
 * @param {UiIconName} name
 * @returns {SVGSVGElement}
 */
export function uiIcon(doc, name) {
  const svg = doc.createElementNS(SVG_NS, "svg");
  for (const [k, v] of [["viewBox", "0 0 24 24"], ["fill", "none"], ["stroke", "currentColor"], ["stroke-width", "2"], ["stroke-linecap", "round"], ["stroke-linejoin", "round"], ["focusable", "false"], ["aria-hidden", "true"], ["class", "ui-icon"]]) svg.setAttribute(k, v);
  for (const shape of /** @type {readonly Shape[]} */ (UI_ICONS[name])) {
    const part = doc.createElementNS(SVG_NS, shape.tag);
    for (const [key, value] of Object.entries(shape)) if (key !== "tag") part.setAttribute(key, String(value));
    svg.appendChild(part);
  }
  return svg;
}
