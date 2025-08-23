/**** Email Attachment Indexer — date basis, quick ranges, quarters, prompts & remembered choices ****/
/* New in this version:
 * - Date basis: Received (message-level) vs Threaded (conversation-level)
 * - Quick ranges: MTD, QTD, YTD, Since last run (UI fills dates)
 * - Year + Quarter picker: enter year, choose any of Q1–Q4 to build an inclusive date range
 * - Still includes: domain/email filter, Drive-save with type checkboxes, destination prompts, clear remembered choices
 */

const LOG_HEADERS = [
  'Logged Key',        // hidden de-dupe key
  'File Name',
  'File Type',         // MIME
  'Size (KB)',
  'Email Subject',
  'From',
  'To',
  'Date',
  'Message Link',
  'Thread Link',
  'Saved File URL'
];

const PREFS_KEY = 'ATT_SIDEBAR_PREFS';
const LAST_RUN_KEY = 'ATT_LAST_RUN_ISO';
const DEST_PREF_MATCH   = 'ATT_DEST_PREF_MATCH';   // 'append' | 'new'
const DEST_PREF_FOREIGN = 'ATT_DEST_PREF_FOREIGN'; // 'new' | 'append_anyway'

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Get Attachments')
    .addItem('Open Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Clear remembered choices', 'clearRememberedChoices')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar');
  const prefs = getPrefs_();
  prefs.lastRunISO = PropertiesService.getUserProperties().getProperty(LAST_RUN_KEY) || '';
  html.prefs = prefs;
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('Email Attachments'));
}

/** ===== Preferences (sidebar fields) ===== */
function getPrefs_() {
  const raw = PropertiesService.getUserProperties().getProperty(PREFS_KEY);
  if (!raw) {
    return {
      dateMode: 'from',        // 'from' | 'range'
      dateBasis: 'received',   // 'received' | 'thread'
      fromDate: '',
      endDate: '',
      domain: '',
      senderEmail: '',
      saveToDrive: false,
      folderUrl: '',
      includeInline: false,
      minSizeKB: 20,
      types: {
        pdf: true, word: false, ppt: false, excel: true, csvtsv: true,
        txt: false, json: false, images: false, archives: false, code: false, other: false
      }
    };
  }
  try {
    const obj = JSON.parse(raw);
    obj.dateMode  = obj.dateMode  || 'from';
    obj.dateBasis = obj.dateBasis || 'received';
    obj.fromDate  = obj.fromDate  || '';
    obj.endDate   = obj.endDate   || '';
    obj.senderEmail = obj.senderEmail || '';
    return obj;
  } catch (_) {
    return {};
  }
}

function setPrefs_(prefs) {
  PropertiesService.getUserProperties().setProperty(PREFS_KEY, JSON.stringify(prefs));
}

/** ===== Public: called from Sidebar to run job ===== */
function runAttachmentJobFromSidebar(form) {
  setPrefs_(form); // remember user selections

  // ====== VALIDATION ======
  const mode = form.dateMode || 'from';
  const basis = (form.dateBasis === 'thread') ? 'thread' : 'received';

  if (!/^\d{4}-\d{2}-\d{2}$/.test(form.fromDate || '')) {
    throw new Error('Please enter Start Date as YYYY-MM-DD, e.g., 2025-05-01.');
  }
  if (mode === 'range') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(form.endDate || '')) {
      throw new Error('Please enter End Date as YYYY-MM-DD, e.g., 2025-07-31.');
    }
    const sd = new Date(form.fromDate + 'T00:00:00');
    const ed = new Date(form.endDate + 'T00:00:00');
    if (ed < sd) throw new Error('End Date must be the same as or after Start Date.');
  }

  const hasDomain = !!(form.domain && form.domain.trim());
  const hasEmail  = !!(form.senderEmail && form.senderEmail.trim());

  if (hasDomain && hasEmail) {
    throw new Error('Please fill EITHER "Sender domain" OR "Specific sender email" (leave the other blank).');
  }
  if (hasDomain && !/^@.+\..+/.test(form.domain.trim())) {
    throw new Error('Sender domain should look like "@xyz.com"');
  }
  if (hasEmail && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(form.senderEmail.trim())) {
    throw new Error('Specific sender email should look like "name@.com"');
  }
  if (form.saveToDrive) {
    if (!form.folderUrl || !extractFolderIdFromUrl_(form.folderUrl)) {
      throw new Error('Please paste a valid Drive FOLDER URL, e.g., https://drive.google.com/drive/folders/XXXXXXXXXXXXXXXX');
    }
  }
  if (form.minSizeKB && isNaN(Number(form.minSizeKB))) {
    throw new Error('Min size (KB) should be a number, e.g., 20');
  }

  // ====== QUERY BUILD ======
  const startStrForQuery = form.fromDate.replaceAll('-', '/');
  let query = `after:${startStrForQuery}`;
  let endBoundary = null;

  if (mode === 'range') {
    const ed = new Date(form.endDate + 'T00:00:00');
    endBoundary = new Date(ed.getTime());
    endBoundary.setDate(endBoundary.getDate() + 1); // inclusive end → before: End+1
    const y = endBoundary.getFullYear();
    const m = String(endBoundary.getMonth() + 1).padStart(2, '0');
    const d = String(endBoundary.getDate()).padStart(2, '0');
    query += ` before:${y}/${m}/${d}`;
  }

  if (hasEmail) {
    query += ` from:${form.senderEmail.trim()}`;
  } else if (hasDomain) {
    query += ` from:*${form.domain.trim()}`;
  }

  // ====== DESTINATION PICK ======
  const dest = pickDestinationSheet_();
  const sheet = dest.sheet;
  const toolManaged = dest.toolManaged; // true = headers match or new tab we created

  // Ensure headers exist on tool-managed sheets
  if (toolManaged && sheet.getLastRow() === 0) {
    sheet.appendRow(LOG_HEADERS);
    try { sheet.setFrozenRows(1); } catch (_) {}
    try { sheet.hideColumns(1); } catch (_) {}
    try { sheet.autoResizeColumns(1, LOG_HEADERS.length); } catch (_) {}
  }

  // Load de-dup keys
  const lastRow = sheet.getLastRow();
  const existingKeys = new Set();
  if (toolManaged && lastRow > 1) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
    keys.forEach(k => existingKeys.add(String(k)));
  }

  // ====== RUN SEARCH ======
  const tz = Session.getScriptTimeZone();
  const startDate = new Date(form.fromDate + 'T00:00:00');
  const PAGE_SIZE = 150;
  let start = 0;
  let processedThreads = 0;
  let processedAttachments = 0;
  const buffer = [];

  const savePred = buildSavePredicate_(form.types);
  const includeInline = !!form.includeInline;
  const minSize = Number(form.minSizeKB || 0);
  const folderId = form.saveToDrive ? extractFolderIdFromUrl_(form.folderUrl) : null;

  // Helpers for date checks
  const isInRangeMsg = (d) => {
    if (d < startDate) return false;
    if (endBoundary && !(d < endBoundary)) return false; // end exclusive
    return true;
  };
  const isInRangeThread = (thread) => {
    const td = thread.getLastMessageDate();
    if (!td) return false;
    return isInRangeMsg(td);
  };

  while (true) {
    const threads = GmailApp.search(query, start, PAGE_SIZE);
    if (!threads || threads.length === 0) break;

    for (let t = 0; t < threads.length; t++) {
      try {
        const thread = threads[t];
        const threadId = thread.getId();
        const threadLink = `https://mail.google.com/mail/u/0/#all/${threadId}`;
        const msgs = thread.getMessages();

        // Thread-level gating if using threaded basis
        if (basis === 'thread') {
          if (!isInRangeThread(thread)) continue;
        }

        for (let mi = 0; mi < msgs.length; mi++) {
          const msg = msgs[mi];

          // Message-level gating only for "received" basis
          if (basis === 'received') {
            if (!isInRangeMsg(msg.getDate())) continue;
          }

          const messageId = msg.getId();
          const messageLink = `https://mail.google.com/mail/u/0/#inbox/${messageId}`;

          const atts = msg.getAttachments({
            includeInlineImages: includeInline,
            includeAttachments: true
          });
          if (!atts || atts.length === 0) continue;

          for (let ai = 0; ai < atts.length; ai++) {
            const att = atts[ai];
            const sizeKB = Math.round(att.getSize() / 1024);
            if (minSize && sizeKB < minSize) continue;

            const key = `${messageId}#${ai}#${att.getSize()}`;
            if (toolManaged && existingKeys.has(key)) continue; // de-dup on tool-managed tabs

            let savedUrl = '';
            if (form.saveToDrive && savePred(att)) {
              const file = DriveApp.getFolderById(folderId).createFile(att);
              file.setDescription(`Saved from Gmail message ${messageId}`);
              savedUrl = file.getUrl();
            }

            buffer.push([
              key,
              att.getName(),
              att.getContentType(),
              sizeKB,
              msg.getSubject() || '',
              msg.getFrom() || '',
              msg.getTo() || '',
              Utilities.formatDate(msg.getDate(), tz, 'yyyy-MM-dd HH:mm'),
              messageLink,
              threadLink,
              savedUrl
            ]);
            processedAttachments++;
          }
        }
        processedThreads++;
      } catch (err) {
        console.warn(`Error on thread index ${start + t}: ${err}`);
      }
    }

    if (buffer.length >= 800) writeBuffer_(sheet, buffer, LOG_HEADERS.length);

    start += threads.length;
    if (threads.length < PAGE_SIZE) break;
  }

  if (buffer.length > 0) writeBuffer_(sheet, buffer, LOG_HEADERS.length);

  if (toolManaged) {
    try { sheet.hideColumns(1); } catch (_) {}
    try { sheet.setFrozenRows(1); } catch (_) {}
    try { sheet.autoResizeColumns(1, LOG_HEADERS.length); } catch (_) {}
  }

  // Save last run timestamp (ISO)
  PropertiesService.getUserProperties().setProperty(LAST_RUN_KEY, new Date().toISOString());

  const dateDesc = (mode === 'range')
    ? `Date filter: ${form.fromDate} to ${form.endDate} (inclusive)`
    : `Date filter: from ${form.fromDate} to now`;

  const basisDesc = (basis === 'thread')
    ? 'Date basis: Threaded (include all attachments from threads active in range)'
    : 'Date basis: Received (only attachments from messages in range)';

  return {
    message:
      `Done.\nThreads processed: ${processedThreads}\nAttachments logged: ${processedAttachments}\n` +
      `Query: ${query}\n${dateDesc}\n${basisDesc}\nDestination: ${sheet.getName()}`
  };
}

/** ===== Destination selection with prompts & remembered choices ===== */
function pickDestinationSheet_() {
  const ui = SpreadsheetApp.getUi();
  const active = SpreadsheetApp.getActiveSheet();

  // Very large sheet safeguard → force new
  const rowCapWarning = active.getMaxRows() > 900000 || active.getLastRow() > 300000;
  if (rowCapWarning) {
    ui.alert('Large sheet detected', 'This sheet is very large. To avoid timeouts, a new tab will be created for this run.', ui.ButtonSet.OK);
    return { sheet: createNewLogSheet_(), toolManaged: true };
  }

  const isBlank = (active.getLastRow() === 0);
  if (isBlank) {
    return { sheet: active, toolManaged: true };
  }

  const headersMatch = headersMatch_(active);

  if (headersMatch) {
    // Remembered choice?
    const saved = PropertiesService.getUserProperties().getProperty(DEST_PREF_MATCH);
    if (saved === 'append') return { sheet: active, toolManaged: true };
    if (saved === 'new')    return { sheet: createNewLogSheet_(), toolManaged: true };

    const n = Math.max(active.getLastRow() - 1, 0);
    const msg = `This sheet looks like an Attachment Log (headers match).\nIt already has ${n} logged rows.\n\nYES = Append here (recommended)\nNO = Create a new tab\nCANCEL = Cancel`;
    const btn = ui.alert('Use this Attachment Log?', msg, ui.ButtonSet.YES_NO_CANCEL);
    if (btn === ui.Button.CANCEL) throw new Error('Cancelled by user.');
    const choice = (btn === ui.Button.YES) ? 'append' : 'new';

    const remember = ui.alert('Remember this choice?', 'Remember this choice next time?', ui.ButtonSet.YES_NO);
    if (remember === ui.Button.YES) {
      PropertiesService.getUserProperties().setProperty(DEST_PREF_MATCH, choice);
    }
    return (choice === 'append')
      ? { sheet: active, toolManaged: true }
      : { sheet: createNewLogSheet_(), toolManaged: true };
  }

  // Foreign sheet (headers don’t match)
  const savedForeign = PropertiesService.getUserProperties().getProperty(DEST_PREF_FOREIGN);
  if (savedForeign === 'new')           return { sheet: createNewLogSheet_(), toolManaged: true };
  if (savedForeign === 'append_anyway') return { sheet: active, toolManaged: false };

  const msg2 = 'This sheet isn’t empty and doesn’t match the expected headers.\nAppending could affect your layout.\n\nYES = Create a new tab (recommended)\nNO = Append anyway\nCANCEL = Cancel';
  const btn2 = ui.alert('This sheet has existing content', msg2, ui.ButtonSet.YES_NO_CANCEL);
  if (btn2 === ui.Button.CANCEL) throw new Error('Cancelled by user.');
  const choice2 = (btn2 === ui.Button.YES) ? 'new' : 'append_anyway';

  const remember2 = ui.alert('Remember this choice?', 'Remember this choice next time?', ui.ButtonSet.YES_NO);
  if (remember2 === ui.Button.YES) {
    PropertiesService.getUserProperties().setProperty(DEST_PREF_FOREIGN, choice2);
  }
  return (choice2 === 'new')
    ? { sheet: createNewLogSheet_(), toolManaged: true }
    : { sheet: active, toolManaged: false };
}

/** Create a new “Email Attachments” tab */
function createNewLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let base = 'Email Attachments';
  let name = base;
  let i = 2;
  while (ss.getSheetByName(name)) {
    name = `${base} (${i++})`;
    if (i > 50) {
      const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
      name = `${base} – ${ts}`;
      break;
    }
  }
  const sh = ss.insertSheet(name);
  sh.appendRow(LOG_HEADERS);
  try { sh.setFrozenRows(1); } catch (_) {}
  try { sh.hideColumns(1); } catch (_) {}
  try { sh.autoResizeColumns(1, LOG_HEADERS.length); } catch (_) {}
  return sh;
}

/** Check if row 1 matches expected headers (case-insensitive, trimmed) */
function headersMatch_(sheet) {
  const width = LOG_HEADERS.length;
  const row1 = sheet.getRange(1, 1, 1, width).getValues()[0];
  const got = row1.map(x => String(x || '').trim().toLowerCase());
  const exp = LOG_HEADERS.map(x => x.toLowerCase());
  for (let i = 0; i < width; i++) {
    if (got[i] !== exp[i]) return false;
  }
  return true;
}

/** ===== Attachment save predicate (checkboxes) ===== */
function buildSavePredicate_(types) {
  const tests = [];

  if (types.pdf) tests.push(m => m === 'application/pdf');

  if (types.word) tests.push(m =>
    m === 'application/msword' ||
    m === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  );

  if (types.ppt) tests.push(m =>
    m === 'application/vnd.ms-powerpoint' ||
    m === 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  );

  if (types.excel) tests.push(m =>
    m === 'application/vnd.ms-excel' ||
    m === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  if (types.csvtsv) tests.push(m =>
    m === 'text/csv' || m === 'text/tab-separated-values'
  );

  if (types.txt) tests.push(m => m === 'text/plain');

  if (types.json) tests.push(m => m === 'application/json');

  if (types.images) tests.push(m => (m || '').startsWith('image/'));

  if (types.archives) tests.push(m =>
    m === 'application/zip' ||
    m === 'application/x-rar-compressed' ||
    m === 'application/x-7z-compressed' ||
    m === 'application/x-tar' ||
    m === 'application/gzip'
  );

  if (types.code) tests.push(m =>
    (m || '').startsWith('text/x-') || m === 'text/javascript' || m === 'application/x-python-code'
  );

  if (types.other) tests.push(_ => true); // catch-all if explicitly chosen

  if (tests.length === 0) return _ => false;
  return att => {
    const m = att.getContentType() || '';
    return tests.some(fn => fn(m));
  };
}

/** ===== Utilities ===== */
function writeBuffer_(sheet, buffer, width) {
  sheet.getRange(sheet.getLastRow() + 1, 1, buffer.length, width).setValues(buffer);
  buffer.length = 0;
}

function extractFolderIdFromUrl_(url) {
  const m1 = url && url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  const m2 = url && url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  return null;
}

/** Menu: Clear remembered choices (destination decisions + sidebar prefs) */
function clearRememberedChoices() {
  const up = PropertiesService.getUserProperties();
  up.deleteProperty(PREFS_KEY);            // clears sidebar last-used form values (incl. date basis/mode)
  up.deleteProperty(DEST_PREF_MATCH);      // clears “headers match” decision
  up.deleteProperty(DEST_PREF_FOREIGN);    // clears “headers don’t match” decision
  SpreadsheetApp.getUi().alert('Remembered choices cleared.\n(Form defaults and destination prompts will show again next time.)');
}
