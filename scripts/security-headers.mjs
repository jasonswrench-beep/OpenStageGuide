// What every page the program serves says to the browser about what it may load and who may show it in a frame.
//
// The program's own pages (program/) are strict: scripts, styles and pictures from this server only, and no inline script or style. The tools'
// page (public/taskpane.html and its neighbours) is the add-in's pane: no inline script either, but it sets some styles in its markup, so those
// are allowed. Both may be shown in a frame only by a page of this server (the tools sit in a frame of the program's page), so no other website
// can put the editor in a hidden frame and click for the person.
export const PROGRAM_CSP="default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-src 'self'; frame-ancestors 'self'; base-uri 'none'; form-action 'none'; object-src 'none'";
export const PANE_CSP="default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; frame-src 'self'; frame-ancestors 'self'; base-uri 'none'; form-action 'none'; object-src 'none'";

/** The headers that go with an HTML page (and with the program's own files). @param {string} csp */
export const pageHeaders=csp=>({'Content-Security-Policy':csp,'X-Frame-Options':'SAMEORIGIN'});

/**
 * The tools' page names Office.js, the add-in's way of talking to Word, from the web. The program has no Word, and a page shown in a frame that
 * loads it is sent away to a blank one, so every copy of the page the program serves leaves that script out.
 * @param {string | Buffer} html
 */
export const stripOfficeScript=html=>String(html).replace(/<script src="https:\/\/appsforoffice\.microsoft\.com[^"]*"><\/script>\s*/g,'');
