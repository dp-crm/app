// ═══════════════════════════════════════════════
// SHARED CONFIG — the single source of truth for the Apps Script URL
// used by all 5 client-facing pages (upload, feedback, gatherinfo,
// review, findoc).
//
// Previously each of those 5 files hardcoded its own independent copy of
// this URL. That's exactly what caused a real outage in this project:
// the deployment URL changed, only some files got updated, and the
// mismatch wasn't caught until every one of the 5 pages was reported as
// "spinning forever, never loads." Centralizing it here means updating
// the URL is ONE edit, in ONE file, instead of hunting through 5 — the
// class of bug that already happened once becomes structurally much
// harder to repeat.
//
// crm.html is deliberately NOT included here — it already has its own,
// more flexible mechanism (a Settings field, changeable by an admin
// without needing a redeploy at all). This file exists specifically for
// the 5 static pages, which have no such settings UI of their own.
//
// TO UPDATE: change the URL below, then push this one file to GitHub —
// all 5 pages pick up the change immediately, no other file needs editing.
// ═══════════════════════════════════════════════
var SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbziQxDrgnJOOXpI7-FPM00pI6I3v_ELBzURAnyG7SSQglFgoSo0sMuPkfhf4TxVmJFG/exec';
