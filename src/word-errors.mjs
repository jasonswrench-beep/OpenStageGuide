// Word reports a failure as a short code ("InvalidArgument", "ItemNotFound"). That tells an author nothing about what to do, so this turns
// each code into a plain sentence about what went wrong and the step that is most likely to fix it. The code itself stays, small, at the
// end of the message, for whoever has to look into it. Errors that are not Word's (the pane's own messages) pass through unchanged.
const REFRESH='Click once in your document, choose Refresh in the pane, then try again.';
const BY_HAND='If it happens again, press Ctrl+Z to take back the last change and make the change by hand in Word.';
const WAIT='Wait until Word has finished what it is doing (look at the status bar at the bottom of the window), click once in your document, then try again.';

const CODES=Object.freeze({
  InvalidArgument:{what:'Word would not accept one of the changes the pane asked for.',fix:`${REFRESH} ${BY_HAND}`},
  ItemNotFound:{what:'Something the pane was working on is no longer in the document. It was probably moved, deleted or changed while the pane was busy.',fix:'Choose Refresh in the pane, put your cursor in the part you want, then try again.'},
  AccessDenied:{what:'Word will not let the pane change this document.',fix:'The document may be read-only or protected. Choose Enable Editing if Word shows a yellow bar at the top, or Review > Restrict Editing > Stop Protection, then try again.'},
  InvalidSelection:{what:'Word could not use the place your cursor is in.',fix:'Click once in the body of the guide (not in a header, a footer, a text box or the contents table), then try again.'},
  GeneralException:{what:'Word could not carry that out just now.',fix:WAIT},
  InvalidOperation:{what:'Word could not carry that out just now.',fix:WAIT},
  InvalidObjectPath:{what:'Word could not find the part of the document the pane meant.',fix:REFRESH},
  NotImplemented:{what:'This version of Word does not support that step.',fix:'Update Word (File > Account > Update Options > Update Now), or make the change by hand in the document.'},
  UnsupportedOperation:{what:'This version of Word does not support that step.',fix:'Update Word (File > Account > Update Options > Update Now), or make the change by hand in the document.'},
  ApiNotFound:{what:'This version of Word does not support that step.',fix:'Update Word (File > Account > Update Options > Update Now), or make the change by hand in the document.'},
  ActivityLimitReached:{what:'That was too much for Word to do in one go.',fix:'Wait a few seconds and try again, or save the document and reopen it first.'},
  RequestPayloadSizeLimitExceeded:{what:'That was too much for Word to do in one go.',fix:'Try it on a smaller part of the guide, or save the document and reopen it first.'},
  ResponsePayloadSizeLimitExceeded:{what:'The document is too big for Word to hand over in one go.',fix:'Save the document and reopen it, then try again.'},
  RequestTimeout:{what:'Word took too long to answer.',fix:WAIT},
  Timeout:{what:'Word took too long to answer.',fix:WAIT}
});
const UNKNOWN={what:'Word reported a problem without saying what it was.',fix:`${REFRESH} ${BY_HAND}`};

// Word's codes are words run together in capitals ("InvalidArgument"), unlike Node's ("ENOENT", "ERR_ASSERTION").
export const isWordError=e=>!!e&&!e.explained&&(e.name==='RichApiError'||(typeof e.code==='string'&&/^([A-Z][a-z]+)+$/.test(e.code)));

export function explainWordError(e){
  if(!isWordError(e))return e;
  const known=CODES[e.code]||UNKNOWN,where=e.debugInfo?.errorLocation;
  const out=new Error(`${known.what} ${known.fix} (Word's code: ${e.code||'none given'}${where?`, in ${where}`:''}.)`);
  out.explained=true;out.code=e.code;out.cause=e;
  return out;
}
