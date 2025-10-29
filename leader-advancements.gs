/**************************************
 * Leader Advancements Self Service
 * Troop 3 (B) + Troop 1003 (G)
 * Google Apps Script (Spreadsheet-bound)
 * Version: v2025-10-22.1 (MID allowlist + sync + 2AM trigger)

 **************************************/


/*** ====== BASIC UTILS & CONFIG ====== ***/


// Spreadsheet & tabs
const SS_ID = SpreadsheetApp.getActive().getId(); // bound to: Leader Advancements Self Service (Responses)
const SHEET_FORM_RESPONSES = 'Form Responses 1';   // auto-created by Forms
const SHEET_LEADERS = 'Leaders';
const SHEET_SUBS = 'Submissions';
const SHEET_DATA = 'Data_Settings';
const SHEET_MID_ALLOW = 'MID_AllowList';          // NEW: local allow-list tab (MID | First | Last)


// Data tabs used to populate dropdowns
const RANGE_AWARDS = 'Data_Awards!A1:B999';                // Code | Name
const RANGE_TRAINING = 'TrainingCatalog!A1:B999';          // Code | Friendly Title (or A1:C999 if you add Popular)
const RANGE_SERVICE_PIN_YEARS = 'ServicePinYears!A1:A999'; // Header in A1, numbers below


// Settings keys (read from Data_Settings)
const SETTINGS = {
  ENV: 'ENV',
  DL_LEADER_ADV: 'DL_LEADER_ADV',
  DL_WEBMASTER: 'DL_WEBMASTER',
  RATE_LIMIT_MAX_HITS: 'RATE_LIMIT_MAX_HITS',
  RATE_LIMIT_WINDOW_MIN: 'RATE_LIMIT_WINDOW_MIN',
  PIN_LOOKAHEAD_DAYS: 'PIN_LOOKAHEAD_DAYS',
  PIN_NUDGE_DAYS: 'PIN_NUDGE_DAYS',
  REMINDER_DEFAULT_DAYS: 'REMINDER_DEFAULT_DAYS',
  SEND_HISTORY_ALWAYS: 'SEND_HISTORY_ALWAYS',
  MID_MASTER_URL: 'MID_MASTER_URL' // NEW: Google Sheet (master roster) URL
};


// Script properties (for ACTION_TOKEN if you deploy action links as a Web App)
const PROP = PropertiesService.getScriptProperties();


// Helpers
function getSheet_(name) { return SpreadsheetApp.openById(SS_ID).getSheetByName(name); }
function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && sh.getLastRow() === 0) sh.appendRow(headers);
  return sh;
}
function getSettings_() {
  const sh = getSheet_(SHEET_DATA);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const kIdx = h.indexOf('Key'), vIdx = h.indexOf('Value');
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const k = (data[i][kIdx] || '').toString().trim();
    const v = (data[i][vIdx] || '').toString().trim();
    if (k) out[k] = v;
  }
  return out;
}


// Environment label (PROD/TEST). Default to PROD unless explicitly set.
function envTag_() { return (getSettings_()[SETTINGS.ENV] || 'PROD').toUpperCase(); }
function subject_(parts) { return parts.filter(Boolean).join(' • '); }
function html_(s){ return String(s ?? '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
function now_(){ return new Date(); }


// Simple rate-limit by MID
function rateLimited_(mid) {
  const s = getSettings_();
  const maxHits = parseInt(s[SETTINGS.RATE_LIMIT_MAX_HITS] || '3', 10);
  const windowMin = parseInt(s[SETTINGS.RATE_LIMIT_WINDOW_MIN] || '5', 10);
  const cache = CacheService.getScriptCache();
  const key = `rl:${mid}`;
  const hits = Number(cache.get(key) || '0') + 1;
  cache.put(key, String(hits), windowMin * 60);
  return hits > maxHits;
}


function labelFooter_(){ return `<br><br><small>Environment: ${envTag_()}</small>`; }
function getLinkedForm_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const formUrl = ss.getFormUrl();
  if (!formUrl) throw new Error('No Google Form is linked to this spreadsheet.');
  return FormApp.openByUrl(formUrl);
}


// Sort/dedupe helper for awards/training (we keep service years ordered separately)
function uniqSorted_(arr) {
  return Array.from(new Set((arr || []).map(s => String(s || '').trim()).filter(Boolean)))
              .sort((a,b) => a.localeCompare(b));
}
function pad2_(v) {
  const n = Number(String(v).trim());
  return Number.isFinite(n) ? (n < 10 ? '0' + n : String(n)) : '';
}


/*** ====== YOUTH KNOT RULES (UNCHANGED) ====== */
const YOUTH_KNOT_CODES = new Set(['REL-YOUTH', 'KNOT-AOL-YOUTH']);
function parseAwardCode_(choiceValue) {
  const s = String(choiceValue || '');
  if (s.includes('—')) return s.split('—')[0].trim();
  return s.split('-')[0].trim();
}
function isYouthOneTimeKnot_(requestedItemLabel) {
  const code = parseAwardCode_(requestedItemLabel);
  return YOUTH_KNOT_CODES.has(code);
}
function hasYouthKnotOnFile_(mid, requestedItemLabel) {
  const code = parseAwardCode_(requestedItemLabel);
  const sh = getOrCreateSheet_(SHEET_SUBS);
  if (sh.getLastRow() < 2) return false;
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const iMID  = h.indexOf('MID');
  const iItem = h.indexOf('Requested Item');
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][iMID]) !== String(mid)) continue;
    if (parseAwardCode_(String(data[r][iItem])) === code) return true;
  }
  return false;
}


/*** ====== Q MAP (EXACT form question titles) ====== */
const Q = {
  // Identity
  EMAIL_ADDR: 'Email Address',               // Forms built-in
  MID: 'Member ID (MID)',
  FIRST: 'First Name',
  LAST: 'Last Name',
  UNIT: 'Unit(s)',
  POSITION: 'Primary Position',              // optional
  REG_SINCE: 'Registered Since',             // Date, optional


  // Request / Log
  REQ_ITEM_REQ: 'Requested Item (Request)',
  REQ_DETAILS: 'Details/Justification (Request)',
  REQ_PROOF: 'Proof URL (optional) (Request)',


  REQ_ITEM_LOG: 'Requested Item (Log)',
  LOG_DATE: 'Date Issued (Self-Reported)',
  LOG_BY: 'Issued By (Self-Reported)',
  LOG_WHERE: 'Where Issued (Self-Reported)',
  LOG_PROOF: 'Proof URL (optional) (Log)',


  // Service Pin
  PIN_YEARS: 'Years of registered service',
  PIN_ATTEST: 'I attest the tenure is accurate',


  // Training
  TRAIN_COURSE: 'Course (search/choose)',
  TRAIN_DATE: 'Completion Date',
  TRAIN_PROOF: 'Proof URL',
  TRAIN_REMIND: 'Remind me before expiry',


  // Review / Extras
  REVIEW_CHECK: 'I’ve reviewed my answers and they’re correct.',
  SEND_HISTORY: 'Email me my history now',
  LOOKUP_FLAG: 'Lookup My History'
};


/*** ====== DROPDOWNS (unchanged) ====== */
function readAwardsList_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const [sheetName] = RANGE_AWARDS.split('!');
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { Logger.log('WARN: Awards sheet not found'); return []; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rng = sh.getRange(1, 1, lastRow, 2).getValues(); // A1:B(lastRow)
  const h = rng[0].map(v => String(v).trim().toLowerCase());
  const iCode = h.indexOf('code'), iName = h.indexOf('name');
  if (iCode === -1 || iName === -1) return [];
  const out = rng.slice(1).map(r => {
    const code = String(r[iCode]||'').trim();
    const name = String(r[iName]||'').trim();
    return (code && name) ? `${code} — ${name}` : '';
  });
  const final = uniqSorted_(out).slice(0, 1000);
  Logger.log(`Awards choices: ${final.length}`);
  return final;
}
function readTrainingList_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const [sheetName] = RANGE_TRAINING.split('!');
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { Logger.log('WARN: Training sheet not found'); return []; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const cols = Math.min(sh.getLastColumn(), /:C\d+$/i.test(RANGE_TRAINING) ? 3 : 2);
  const rng = sh.getRange(1, 1, lastRow, cols).getValues();
  const h = rng[0].map(v => String(v).trim().toLowerCase());
  const iCode = h.indexOf('code');
  const iTitle = h.indexOf('friendly title');
  const iPopular = h.indexOf('popular'); // -1 if absent
  if (iCode === -1 || iTitle === -1) return [];
  const seen = new Set();
  const rows = [];
  rng.slice(1).forEach(r => {
    const code  = String(r[iCode]  || '').trim();
    const title = String(r[iTitle] || '').trim();
    if (!code || !title) return;
    if (iPopular !== -1 && String(r[iPopular] || '').toUpperCase() !== 'Y') return;
    if (seen.has(code)) return;
    seen.add(code);
    rows.push({ code, title });
  });
  rows.sort((a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }));
  const final = rows.map(x => `${x.code} — ${x.title}`).slice(0, 1000);
  Logger.log(`Training choices (title-sorted): ${final.length}`);
  return final;
}
function readServiceYears_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const [sheetName] = RANGE_SERVICE_PIN_YEARS.split('!');
  const sh = ss.getSheetByName(sheetName);
  if (!sh) { Logger.log('WARN: ServicePinYears sheet not found'); return []; }
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rng = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // A2:A(last)
  const seen = new Set();
  const out = [];
  rng.forEach(r => {
    const raw = String(r[0] ?? '').trim();
    if (!raw) return;
    if (!/^\d+$/.test(raw)) return;
    const val = pad2_(raw);
    if (val && !seen.has(val)) { seen.add(val); out.push(val); }
  });
  Logger.log(`Service years choices (zero-padded): ${out.length}`);
  return out; // preserve sheet order
}
function setDropdownChoices_(form, questionTitle, choices) {
  const vals = uniqSorted_(choices).slice(0, 1000);
  if (!vals.length) { Logger.log(`WARN: Skipping "${questionTitle}" — no choices found.`); return; }
  const item = form.getItems().find(it => it.getTitle().trim() === questionTitle.trim());
  if (!item) throw new Error(`Form item not found: ${questionTitle}`);
  const t = item.getType();
  if (t === FormApp.ItemType.LIST)                 item.asListItem().setChoiceValues(vals);
  else if (t === FormApp.ItemType.MULTIPLE_CHOICE) item.asMultipleChoiceItem().setChoiceValues(vals);
  else if (t === FormApp.ItemType.CHECKBOX)        item.asCheckboxItem().setChoiceValues(vals);
  else throw new Error(`Item "${questionTitle}" is not a list-like question (type=${t})`);
}
function setDropdownChoicesPreserveOrder_(form, questionTitle, choices) {
  const vals = (choices || []).map(s => String(s).trim()).filter(Boolean).slice(0, 1000);
  if (!vals.length) { Logger.log(`WARN: Skipping "${questionTitle}" — no choices found.`); return; }
  const item = form.getItems().find(it => it.getTitle().trim() === questionTitle.trim());
  if (!item) throw new Error(`Form item not found: ${questionTitle}`);
  const t = item.getType();
  if (t === FormApp.ItemType.LIST)                 item.asListItem().setChoiceValues(vals);
  else if (t === FormApp.ItemType.MULTIPLE_CHOICE) item.asMultipleChoiceItem().setChoiceValues(vals);
  else if (t === FormApp.ItemType.CHECKBOX)        item.asCheckboxItem().setChoiceValues(vals);
  else throw new Error(`Item "${questionTitle}" is not a list-like question (type=${t})`);
}


/*** ====== SYNC MENU ACTION ====== */
function syncAllDropdowns() {
  const form = getLinkedForm_();
  setDropdownChoices_(form, Q.REQ_ITEM_REQ, readAwardsList_());
  setDropdownChoices_(form, Q.REQ_ITEM_LOG, readAwardsList_());
  setDropdownChoicesPreserveOrder_(form, Q.TRAIN_COURSE, readTrainingList_());
  setDropdownChoicesPreserveOrder_(form, Q.PIN_YEARS, readServiceYears_());
}


/*** ====== MID ALLOW-LIST: sync + check ====== */


// Manual sync from master Google Sheet URL
function syncMidAllowlist() {
  const s = getSettings_();
  const url = (s[SETTINGS.MID_MASTER_URL] || '').trim();
  if (!/^https:\/\/docs\.google\.com\/spreadsheets\//i.test(url)) {
    throw new Error('Data_Settings → MID_MASTER_URL is missing or not a Google Sheet URL.');
  }


  const master = SpreadsheetApp.openByUrl(url);
  // Try a sheet named "MID_Master"; else, use the first sheet
  let msh = master.getSheetByName('MID_Master') || master.getSheets()[0];
  const vals = msh.getDataRange().getValues();


  // Expect MID in Col A, First in Col C, Last in Col E (as you specified)
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const mid = String(vals[i][0] || '').trim();
    const first = String(vals[i][2] || '').trim();
    const last  = String(vals[i][4] || '').trim();
    if (!/^\d{5,12}$/.test(mid)) continue;
    out.push([mid, first, last]);
  }


  // Dedupe by MID, keep first seen name
  const seen = new Set();
  const cleaned = [];
  out.forEach(r => {
    if (seen.has(r[0])) return;
    seen.add(r[0]);
    cleaned.push(r);
  });


  const sh = getOrCreateSheet_(SHEET_MID_ALLOW, ['MID','First','Last']);
  // Clear and rewrite (keep header)
  if (sh.getLastRow() > 1) sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).clearContent();
  if (cleaned.length) sh.getRange(2,1, cleaned.length, 3).setValues(cleaned);


  SpreadsheetApp.getActive().toast(`Synced ${cleaned.length} MIDs to ${SHEET_MID_ALLOW}.`, 'Scouter Portal', 5);
}


function isMidAllowed_(mid) {
  if (!/^\d{5,12}$/.test(mid)) return false;
  const sh = getOrCreateSheet_(SHEET_MID_ALLOW, ['MID','First','Last']);
  const last = sh.getLastRow();
  if (last < 2) return false;
  const rng = sh.getRange(2,1, last-1, 1).getValues();
  for (let i=0;i<rng.length;i++) {
    if (String(rng[i][0]) === String(mid)) return true;
  }
  return false;
}


// Install a daily sync at 02:00 (script timezone)
function installMidAllowlist2am_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction && t.getHandlerFunction() === 'syncMidAllowlist')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncMidAllowlist').timeBased().atHour(2).everyDays(1).create();
  SpreadsheetApp.getActive().toast('Installed daily MID allow-list sync @ 02:00.', 'Scouter Portal', 5);
}


/*** ====== ON FORM SUBMIT WORKFLOW (with MID enforcement) ====== */


function onFormSubmit(e) {
  const s = getSettings_();


  const nv = e.namedValues;
  const ts = new Date(e.range.getValues()[0][0] || now_());
  const email = (nv[Q.EMAIL_ADDR] || [''])[0].trim();
  const mid   = (nv[Q.MID]        || [''])[0].trim();
  const first = (nv[Q.FIRST]      || [''])[0].trim();
  const last  = (nv[Q.LAST]       || [''])[0].trim();
  const unit  = (nv[Q.UNIT]       || [''])[0].trim();
  const position = (nv[Q.POSITION] || [''])[0].trim();
  const regSince = (nv[Q.REG_SINCE] || [''])[0].trim();


  // Quick MID format guard
  if (!/^\d{5,12}$/.test(mid)) {
    MailApp.sendEmail({
      to: email,
      subject: subject_(['Scouter Portal', 'Invalid Member ID format']),
      htmlBody: `We couldn’t process your submission because the Member ID format looks off. Please double-check and resubmit.` + labelFooter_()
    });
    return;
  }


  if (rateLimited_(mid)) {
    MailApp.sendEmail({
      to: email,
      subject: subject_(['Scouter Portal', 'Please wait a moment']),
      htmlBody: `We received multiple submissions in a short time. Please wait a few minutes and try again.` + labelFooter_()
    });
    return;
  }


  // === Enforce MID allow-list (leaders) ===
  if (!isMidAllowed_(mid)) {
    // Log to Submissions for audit
    ensureSubmissionsHeader_();
    const shSubs = getOrCreateSheet_(SHEET_SUBS);
    shSubs.appendRow([
      ts, envTag_(), mid, email, `${first} ${last}`, unit,
      'Invalid MID', '', '', '', 'Invalid MID', '', '', 'No', '', '', '', '', ''
    ]);


    // Blunt, helpful notice to submitter (no fluff)
    MailApp.sendEmail({
      to: email,
      subject: subject_(['Scouter Portal', 'We could not match your Member ID']),
      htmlBody:
        `<p style="color:#b00020;font-weight:bold;margin:0 0 .5rem 0">We couldn’t match the Member ID you entered to our roster.</p>
         <p>Please contact your Committee Chair or an ASM to confirm your MID and have it added, then resubmit.</p>` + labelFooter_()
    });


    // Alert webmaster only (do NOT spam the DL)
    const masterUrl = s[SETTINGS.MID_MASTER_URL] || '#';
    const body = [
      '<b>Invalid/unknown MID submission</b>',
      `<b>Name:</b> ${html_(first)} ${html_(last)} &nbsp; <b>MID:</b> ${html_(mid)}`,
      `<b>Unit(s):</b> ${html_(unit)}`,
      `<p><a href="${masterUrl}" target="_blank" rel="noopener">Open MID Master (update allowlist)</a></p>`
    ].join('<br>');
    MailApp.sendEmail({
      to: s[SETTINGS.DL_WEBMASTER],
      subject: subject_(['Scouter Portal', 'MID not in allowlist']),
      htmlBody: body + labelFooter_()
    });


    return; // stop processing
  }


  // Determine Submission Type by which fields are present
  const hasReq   = !!(nv[Q.REQ_ITEM_REQ] && nv[Q.REQ_ITEM_REQ][0]);
  const hasLog   = !!(nv[Q.REQ_ITEM_LOG] && nv[Q.REQ_ITEM_LOG][0]);
  const hasPin   = !!(nv[Q.PIN_YEARS]    && nv[Q.PIN_YEARS][0]);
  const hasTrain = !!(nv[Q.TRAIN_COURSE] && nv[Q.TRAIN_COURSE][0]);
  const lookupSelected = !!(nv[Q.LOOKUP_FLAG] && nv[Q.LOOKUP_FLAG][0]);


  let submissionType = '';
  if (hasReq) submissionType = 'Request Award / Knot';
  else if (hasLog) submissionType = 'Log Previously Issued Award / Knot';
  else if (hasPin) submissionType = 'Request Service Pin';
  else if (hasTrain) submissionType = 'Report Training Completion';
  else if (lookupSelected) submissionType = 'Lookup My History';


  // Upsert Leaders (don’t overwrite optional blanks)
  upsertLeader_({ mid, first, last, email, unit, position, regSince, ts });


  if (submissionType === 'Lookup My History') {
    emailHistoryForMid_(email, mid);
    notifyDL_({ ts, mid, first, last, email, unit, submissionType, requestedItem: 'History lookup', details: '', proof: '' });
    return;
  }


  // Build Submissions row
  const requestedItem = hasReq ? (nv[Q.REQ_ITEM_REQ][0] || '') :
                       hasLog ? (nv[Q.REQ_ITEM_LOG][0] || '') : '';
  const details = hasReq ? (nv[Q.REQ_DETAILS]?.[0] || '') : '';
  const proof = (hasReq ? (nv[Q.REQ_PROOF]?.[0] || '') :
                 hasLog ? (nv[Q.LOG_PROOF]?.[0] || '') :
                 hasTrain ? (nv[Q.TRAIN_PROOF]?.[0] || '') : '');


  // Youth knot guard
  if (submissionType === 'Request Award / Knot' && requestedItem && isYouthOneTimeKnot_(requestedItem)) {
    if (hasYouthKnotOnFile_(mid, requestedItem)) {
      MailApp.sendEmail({
        to: email,
        subject: subject_(['Already on file', 'Youth knot', `[MID ${mid}]`]),
        htmlBody:
          `Hi ${html_(first)},<br><br>` +
          `Our records already show <b>${html_(requestedItem)}</b> on file for your adult record. ` +
          `If something needs correction, reply and we’ll fix it—no need to resubmit.` + labelFooter_()
      });
      notifyDL_({
        ts, mid, first, last, email, unit,
        submissionType: 'Duplicate youth knot attempt',
        requestedItem, details: '', proof: ''
      });
      return;
    }
  }


  const shSubs = getOrCreateSheet_(SHEET_SUBS); // ensure exists
  ensureSubmissionsHeader_();


  const baseRow = [
    ts, envTag_(), mid, email, `${first} ${last}`, unit, submissionType,
    requestedItem, details, proof, '', '', '', 'No', '', '', '', '', ''
  ];


  let status = 'Pending';
  let extraCols = {};


  if (submissionType === 'Log Previously Issued Award / Knot') {
    status = 'Logged';
    extraCols['Recognition Type'] = 'Log existing';
    extraCols['Date Issued (Self-Reported)'] = (nv[Q.LOG_DATE]?.[0] || '');
    extraCols['Issued By (Self-Reported)'] = (nv[Q.LOG_BY]?.[0] || '');
    extraCols['Where Issued (Self-Reported)'] = (nv[Q.LOG_WHERE]?.[0] || '');
    extraCols['Validation Status'] = 'Auto';
    extraCols['De-dup Check'] = dedupNote_(mid, requestedItem, nv[Q.LOG_DATE]?.[0]);
  }


  if (submissionType === 'Request Service Pin') {
    const regDate = lookupLeaderDate_(getOrCreateSheet_(SHEET_LEADERS), mid);
    const attest = (nv[Q.PIN_ATTEST]?.[0] || '');
    extraCols['I attest the tenure is accurate'] = attest;
    if (!regDate) status = 'More Info';
  }


  shSubs.appendRow(baseRow);
  const lastRow = shSubs.getLastRow();
  setSubsColumns_(lastRow, { 'Status': status, ...extraCols });


  const ctx = { ts, mid, first, last, email, unit, submissionType, requestedItem, details, proof };
  notifyDL_(ctx);
  sendReceipt_(ctx);


  // History emails (optional)
  const always = (getSettings_()[SETTINGS.SEND_HISTORY_ALWAYS] || '').toUpperCase() === 'Y';
  let wantsHistory = false;
  try {
    const v = (nv[Q.SEND_HISTORY] || [''])[0];
    wantsHistory = String(v).toLowerCase().indexOf('yes') !== -1;
  } catch (_) {}
  if (always || wantsHistory) emailHistoryForMid_(email, mid);
}


/*** ====== HELPERS (Leaders / Submissions) ====== */
function nightlySyncDropdowns() { syncAllDropdowns(); }


function installMidnightSyncTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'nightlySyncDropdowns')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('nightlySyncDropdowns').timeBased().atHour(0).everyDays(1).create();
}


function ensureSubmissionsHeader_() {
  getOrCreateSheet_(SHEET_SUBS, [
    'Timestamp','Environment','MID','Email','Name','Unit(s)','Submission Type','Requested Item',
    'Details/Justification','Proof URL','Status','Reviewer','Review Date','Issued?','Issued By',
    'Issued Date','Cost','Inventory Source','Notes',
    'Recognition Type','Date Issued (Self-Reported)','Issued By (Self-Reported)','Where Issued (Self-Reported)',
    'Validation Status','De-dup Check'
  ]);
}
function upsertLeader_(obj) {
  const sh = getOrCreateSheet_(SHEET_LEADERS, [
    'MID','First Name','Last Name','Email','Unit(s)','Primary Position',
    'Registered Since (YYYY-MM-DD)','Years of Service','Recommended Service Pin',
    'Last Pin Issued (Years)','Last Submission','Admin Notes'
  ]);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idx = {
    MID: h.indexOf('MID'), FIRST: h.indexOf('First Name'), LAST: h.indexOf('Last Name'),
    EMAIL: h.indexOf('Email'), UNIT: h.indexOf('Unit(s)'), POS: h.indexOf('Primary Position'),
    REG: h.indexOf('Registered Since (YYYY-MM-DD)'), LAST_SUB: h.indexOf('Last Submission')
  };
  const map = {};
  for (let i=1;i<data.length;i++) map[String(data[i][idx.MID])] = i+1;


  const existingRow = map[obj.mid];
  if (existingRow) {
    if (obj.first) sh.getRange(existingRow, idx.FIRST+1).setValue(obj.first);
    if (obj.last) sh.getRange(existingRow, idx.LAST+1).setValue(obj.last);
    if (obj.email) sh.getRange(existingRow, idx.EMAIL+1).setValue(obj.email);
    if (obj.unit) sh.getRange(existingRow, idx.UNIT+1).setValue(obj.unit);
    if (obj.position) sh.getRange(existingRow, idx.POS+1).setValue(obj.position);
    if (obj.regSince) {
      const d = normalizeDate_(obj.regSince);
      sh.getRange(existingRow, idx.REG+1).setValue(d ? d : obj.regSince);
    }
    sh.getRange(existingRow, idx.LAST_SUB+1).setValue(obj.ts);
  } else {
    const d = obj.regSince ? normalizeDate_(obj.regSince) : '';
    sh.appendRow([obj.mid,obj.first,obj.last,obj.email,obj.unit,obj.position,d,'','','',obj.ts,'']);
  }
}
function normalizeDate_(v) {
  try {
    if (Object.prototype.toString.call(v) === '[object Date]') return v;
    if (!v) return null;
    const d = new Date(v);
    return isNaN(d) ? null : d;
  } catch (_) { return null; }
}
function lookupLeaderDate_(leadersSheet, mid) {
  const data = leadersSheet.getDataRange().getValues();
  const h = data[0];
  const iMid = h.indexOf('MID'), iReg = h.indexOf('Registered Since (YYYY-MM-DD)');
  for (let i=1;i<data.length;i++) {
    if (String(data[i][iMid]) === String(mid)) return data[i][iReg] || '';
  }
  return '';
}
function setSubsColumns_(row, map) {
  const sh = getOrCreateSheet_(SHEET_SUBS);
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  Object.keys(map).forEach(name => {
    const idx = h.indexOf(name);
    if (idx !== -1) sh.getRange(row, idx+1).setValue(map[name]);
  });
}
function dedupNote_(mid, item, dateStr) {
  const sh = getOrCreateSheet_(SHEET_SUBS);
  if (sh.getLastRow() < 2) return '';
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const iMID = h.indexOf('MID'), iItem = h.indexOf('Requested Item'), iDate = h.indexOf('Date Issued (Self-Reported)');
  const newDate = dateStr ? new Date(dateStr) : null;
  for (let i=1;i<data.length;i++) {
    if (String(data[i][iMID]) !== String(mid)) continue;
    if (String(data[i][iItem]) !== String(item)) continue;
    if (!newDate || !data[i][iDate]) continue;
    const d = new Date(data[i][iDate]);
    const diff = Math.abs((newDate - d) / (1000*60*60*24));
    if (diff <= 365) return `Possible duplicate of entry dated ${d.toISOString().slice(0,10)}`;
  }
  return '';
}


/*** ====== NOTIFICATIONS ====== */
function notifyDL_(ctx) {
  const s = getSettings_();
  const sheetUrl = SpreadsheetApp.openById(SS_ID).getUrl();
  const body = `
<b>New submission</b><br>
<b>When:</b> ${html_(ctx.ts)}<br>
<b>Name:</b> ${html_(ctx.first)} ${html_(ctx.last)} (${html_(ctx.email)})<br>
<b>Unit:</b> ${html_(ctx.unit)}<br>
<b>MID:</b> ${html_(ctx.mid)}<br>
<b>Type:</b> ${html_(ctx.submissionType)}<br>
<b>Item:</b> ${html_(ctx.requestedItem)}<br>
<b>Details:</b> ${html_(ctx.details)}<br>
<b>Proof:</b> ${html_(ctx.proof)}<br><br>
<a href="${sheetUrl}">Open Submissions Sheet</a>
${labelFooter_()}
`;
  MailApp.sendEmail({
    to: s[SETTINGS.DL_LEADER_ADV],
    subject: subject_(['Scouter Portal', `[MID ${ctx.mid}]`, ctx.submissionType, ctx.requestedItem]),
    htmlBody: body
  });
}
function sendReceipt_(ctx) {
  const msg = `Thanks, ${html_(ctx.first)}. We logged your submission: <b>${html_(ctx.submissionType)} — ${html_(ctx.requestedItem)}</b>.` + labelFooter_();
  MailApp.sendEmail({ to: ctx.email, subject: subject_(['We received your submission']), htmlBody: msg });
}


/*** ====== HISTORY EMAIL ====== */
function emailHistoryForMid_(to, mid) {
  const sh = getOrCreateSheet_(SHEET_SUBS);
  ensureSubmissionsHeader_();
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const rows = data.slice(1).filter(r => String(r[h.indexOf('MID')]) === String(mid));
  const html = rows.length
    ? '<table border="1" cellpadding="6" cellspacing="0"><tr>' + h.map(x=>`<th>${html_(x)}</th>`).join('') + '</tr>' +
      rows.map(r=>'<tr>'+r.map(c=>`<td>${html_(c)}</td>`).join('')+'</tr>').join('') + '</table>'
    : 'No prior submissions found for this MID.';
  MailApp.sendEmail({ to, subject: subject_(['Your Scouter history', `[MID ${mid}]`]), htmlBody: html + labelFooter_() });
}


/*** ====== ACTION LINKS (Approve / Deny / Need Info / Mark Issued) ====== */
function actionLinks_(mid, item) {
  const webUrl = ScriptApp.getService().getUrl(); // set after Web App deploy
  const token = PROP.getProperty('ACTION_TOKEN') || 'SET_ME';
  function link(a,label){ return `<a href="${webUrl}?a=${encodeURIComponent(a)}&mid=${encodeURIComponent(mid)}&item=${encodeURIComponent(item)}&t=${encodeURIComponent(token)}" target="_blank" rel="noopener">${label}</a>`; }
  return ['<b>Quick Actions:</b>', link('Approve','Approve'), link('Deny','Deny'), link('MoreInfo','Need Info'), link('MarkIssued','Mark Issued')].join(' | ');
}
function doGet(e) {
  const p = e.parameter || {};
  if ((p.page || '') === 'history') return ContentService.createTextOutput('History page not enabled.');
  const token = p.t || '';
  const ok = token && token === PROP.getProperty('ACTION_TOKEN');
  if (!ok) return ContentService.createTextOutput('Invalid token.');


  const a = p.a || '', mid = p.mid || '', item = p.item || '';
  const sh = getOrCreateSheet_(SHEET_SUBS);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idx = name => h.indexOf(name);
  const iMID = idx('MID'), iItem = idx('Requested Item');


  let row = -1;
  for (let r = data.length - 1; r >= 1; r--) {
    if (String(data[r][iMID]) === String(mid) && String(data[r][iItem]) === String(item)) { row = r + 1; break; }
  }
  if (row === -1) return ContentService.createTextOutput('Submission not found.');


  const me = Session.getActiveUser().getEmail() || 'someone@example.org';
  const set = (col,val)=>sh.getRange(row, idx(col)+1).setValue(val);


  switch (a) {
    case 'Approve':   set('Status','Approved'); set('Reviewer',me); set('Review Date', new Date()); break;
    case 'Deny':      set('Status','Denied');   set('Reviewer',me); set('Review Date', new Date()); break;
    case 'MoreInfo':  set('Status','More Info');set('Reviewer',me); set('Review Date', new Date()); break;
    case 'MarkIssued':set('Issued?','Yes');     set('Issued By',me); set('Issued Date', new Date());
                      if (!sh.getRange(row, idx('Status')+1).getValue()) set('Status','Approved'); break;
    default:          return ContentService.createTextOutput('Unknown action.');
  }


  const s = getSettings_();
  const sub = subject_(['Scouter Portal', `[MID ${mid}]`, a, item]);
  const body = `Action recorded: <b>${html_(a)}</b> for MID ${html_(mid)} — ${html_(item)}.${labelFooter_()}`;
  MailApp.sendEmail({ to: s[SETTINGS.DL_LEADER_ADV], subject: sub, htmlBody: body });


  return ContentService.createHtmlOutput(`<p>Recorded: <b>${html_(a)}</b> for MID ${html_(mid)} — ${html_(item)}.</p>`);
}


/*** ====== MENU & REMINDERS ====== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scouter Awards')
    .addItem('Issue & Email (selected row)', 'issueAndEmailSelected')
    .addSeparator()
    .addItem('Sync dropdowns from Data tab', 'syncAllDropdowns')
    .addItem('Email history by MID…', 'emailHistoryPrompt_')
    .addSeparator()
    .addItem('Sync MID allowlist (from master)', 'syncMidAllowlist')      // NEW
    .addItem('Install 2:00 AM MID sync', 'installMidAllowlist2am_')       // NEW
    .addToUi();
}


function issueAndEmailSelected() {
  const sh = getOrCreateSheet_(SHEET_SUBS);
  const r = sh.getActiveRange();
  if (!r || r.getRow() < 2) { SpreadsheetApp.getUi().alert('Select a single data row in Submissions.'); return; }
  const row = r.getRow();
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const v = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  const idx = name => h.indexOf(name);


  const email = v[idx('Email')], name = v[idx('Name')], mid = v[idx('MID')], item = v[idx('Requested Item')];


  sh.getRange(row, idx('Status')+1).setValue('Approved');
  sh.getRange(row, idx('Issued?')+1).setValue('Yes');
  sh.getRange(row, idx('Issued By')+1).setValue(Session.getActiveUser().getEmail());
  sh.getRange(row, idx('Issued Date')+1).setValue(new Date());
  sh.getRange(row, idx('Review Date')+1).setValue(new Date());
  sh.getRange(row, idx('Reviewer')+1).setValue(Session.getActiveUser().getEmail());


  const s = getSettings_();
  const sub = subject_(['Recognition Issued', `[MID ${mid}]`, item]);
  const body = `Hi ${html_(name)},<br><br>Your recognition has been issued: <b>${html_(item)}</b>.<br>Congratulations and thank you for your service!` + labelFooter_();
  MailApp.sendEmail({ to: email, cc: s[SETTINGS.DL_LEADER_ADV], subject: sub, htmlBody: body });
}


function emailHistoryPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const midResp = ui.prompt('Email history', 'Enter Member ID (MID):', ui.ButtonSet.OK_CANCEL);
  if (midResp.getSelectedButton() !== ui.Button.OK) return;
  const mid = (midResp.getResponseText() || '').trim();
  if (!/^\d{5,12}$/.test(mid)) { ui.alert('Please enter a valid MID (5–12 digits).'); return; }


  const emailResp = ui.prompt('Email history', 'Enter recipient email:', ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const email = (emailResp.getResponseText() || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) { ui.alert('Please enter a valid email address.'); return; }


  emailHistoryForMid_(email, mid);
  ui.alert('If records exist, the history email has been sent.');
}


/*** ====== TRAINING REMINDERS (unchanged) ====== */
function nightlyTrainingReminders() {
  const s = getSettings_();
  const sh = getOrCreateSheet_(SHEET_SUBS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  const h = data[0];
  const idx = name => h.indexOf(name);


  for (let r=1;r<data.length;r++) {
    if (String(sh.getRange(r+1, idx('Submission Type')+1).getValue()) !== 'Report Training Completion') continue;
    const email = sh.getRange(r+1, idx('Email')+1).getValue();
    const name = sh.getRange(r+1, idx('Name')+1).getValue();
    const item = sh.getRange(r+1, idx('Requested Item')+1).getValue();
    const notes = String(sh.getRange(r+1, idx('Notes')+1).getValue() || '');
    const m = notes.match(/Expires On:\s*(\d{4}-\d{2}-\d{2})/i);
    if (!m) continue;
    const exp = new Date(m[1]+'T00:00:00');
    const daysLeft = Math.floor((exp - new Date())/(1000*60*60*24));
    if ([60,30,7].includes(daysLeft)) {
      const sub = subject_(['Training expiring', item, m[1]]);
      const body = `Hi ${html_(name)},<br><br>Your training <b>${html_(item)}</b> is set to expire on <b>${m[1]}</b>.<br>Please schedule renewal and reply when complete.${labelFooter_()}`;
      MailApp.sendEmail({ to: email, cc: s[SETTINGS.DL_LEADER_ADV], subject: sub, htmlBody: body });
    }
  }
}


// Pin digest & nudges
function monthlyPinDigest() {
  const s = getSettings_();
  const lookahead = parseInt(s[SETTINGS.PIN_LOOKAHEAD_DAYS]||'45',10);
  const sh = getOrCreateSheet_(SHEET_LEADERS);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const i = name => h.indexOf(name);
  const iMID=i('MID'), iF=i('First Name'), iL=i('Last Name'), iE=i('Email'), iU=i('Unit(s)'), iReg=i('Registered Since (YYYY-MM-DD)'), iY=i('Years of Service'), iLast=i('Last Pin Issued (Years)');


  const today = new Date();
  const upcoming = [], overdue = [];


  for (let r=1;r<data.length;r++) {
    const row = data[r];
    const reg = row[iReg] ? new Date(row[iReg]) : null;
    if (!reg || isNaN(reg)) continue;
    const years = Number(row[iY]||0);
    const lastIssued = Number(row[iLast]||0);


    let ann = new Date(reg);
    ann.setFullYear(today.getFullYear());
    if (ann < today) ann.setFullYear(today.getFullYear()+1);


    const targetPin = Math.max(1, years+1);
    if (lastIssued >= targetPin) continue;


    const daysUntil = Math.floor((ann - today)/(1000*60*60*24));
    const obj = {mid: row[iMID], name: row[iF]+' '+row[iL], email: row[iE], unit: row[iU], pin: targetPin, ann};
    if (daysUntil >= 0 && daysUntil <= lookahead) upcoming.push(obj);
    if (ann < today) overdue.push(obj);
  }


  const lines = arr => arr.map(x=>`${html_(x.name)} • MID ${html_(x.mid)} • ${html_(x.unit)} • ${x.pin}-year • Anniversary ${x.ann.toISOString().slice(0,10)}`).join('<br>') || 'None';
  const html = `<b>Service Pins due in next ${lookahead} days:</b><br>${lines(upcoming)}<br><br><b>Overdue Service Pins:</b><br>${lines(overdue)}${labelFooter_()}`;
  MailApp.sendEmail({ to: s[SETTINGS.DL_LEADER_ADV], subject: subject_([`Service Pins Digest (next ${lookahead} days)`]), htmlBody: html });
}


function nightlyPinNudges() {
  const s = getSettings_();
  const nudgeDays = parseInt(s[SETTINGS.PIN_NUDGE_DAYS]||'21',10);
  const sh = getOrCreateSheet_(SHEET_LEADERS);
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const i = name => h.indexOf(name);
  const iMID=i('MID'), iF=i('First Name'), iE=i('Email'), iReg=i('Registered Since (YYYY-MM-DD)'), iY=i('Years of Service'), iLast=i('Last Pin Issued (Years)');


  const today = new Date();
  for (let r=1;r<data.length;r++) {
    const row = data[r];
    const reg = row[iReg] ? new Date(row[iReg]) : null;
    if (!reg || isNaN(reg)) continue;
    const years = Number(row[iY]||0);
    const lastIssued = Number(row[iLast]||0);


    let ann = new Date(reg);
    ann.setFullYear(today.getFullYear());
    if (ann < today) ann.setFullYear(today.getFullYear()+1);


    const targetPin = Math.max(1, years+1);
    const daysUntil = Math.floor((ann - today)/(1000*60*60*24));
    if (daysUntil === nudgeDays && lastIssued < targetPin) {
      const to = row[iE];
      const sub = subject_(['Service Pin upcoming', `[MID ${row[iMID]}]`, `${targetPin}-year`]);
      const body = `Hi ${html_(row[iF])},<br><br>
Looks like your <b>${targetPin}-year Service Pin</b> anniversary is coming up on <b>${ann.toISOString().slice(0,10)}</b>.
We’ll plan to handle recognition at the next court of honor. If your tenure date looks off, reply and we’ll help reconcile it.<br><br>
— Scouter Advancements Team${labelFooter_()}`;
      MailApp.sendEmail({ to, cc: s[SETTINGS.DL_LEADER_ADV], subject: sub, htmlBody: body });
    }
  }
}





