/*
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Sets up a daily trigger to periodically run
 */
function setupDailyTrigger() {
  // Remove existing triggers to prevent duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'fetchDailyEmails') ScriptApp.deleteTrigger(t);
  });

  // Create trigger to run to periodically run
  ScriptApp.newTrigger('fetchDailyEmails')
    .timeBased()
    .everyDays(1)
    .atHour(21)       // 24-hour clock
    .nearMinute(40)
    .create();
}

/**
 * Main function: automatically retrieves today's emails
 */
function fetchDailyEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, 'DailyEmails');

  // ----- 1. Build date range -------------------------------------------------
  const now       = new Date();
  const afterDate = new Date(now);
  afterDate.setDate(afterDate.getDate() - 1);        // yesterday → today

  const tz        = Session.getScriptTimeZone();
  const afterStr  = Utilities.formatDate(afterDate, tz, 'yyyy/MM/dd');
  const beforeStr = Utilities.formatDate(now,      tz, 'yyyy/MM/dd');
  const query     = `after:${afterStr} before:${beforeStr}`;

  // ----- 2. Header row (row 2) ------------------------------------------------
  const headers = [
    'Thread ID', 'Date', 'From Name', 'Subject', 'Message ID',
    'From Email', 'To', 'CC', 'Labels', 'IsUnread',
    'IsStarred', 'IsImportant', 'AttachmentNames', 'Link'
  ];
  if (sheet.getLastRow() < 2 || sheet.getRange(2, 1, 1, headers.length).isBlank()) {
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  }

  // ----- 3. Collect e-mail data ----------------------------------------------
  const data = [];
  GmailApp.search(query).forEach(thread => {
    const threadLabels = thread.getLabels().map(l => l.getName()).join(', ');
    const important    = thread.isImportant();

    thread.getMessages().forEach(msg => {
      const msgDate = msg.getDate();
      if (msgDate < afterDate || msgDate > now) return;      // outside range

      const fromParsed      = parseAddressList(msg.getFrom());
      const attachmentNames = msg.getAttachments().map(a => a.getName()).join(', ');

      data.push([
        thread.getId(),
        msgDate,
        fromParsed.names,
        msg.getSubject(),
        msg.getId(),
        fromParsed.emails,
        msg.getTo() || '',
        msg.getCc() || '',
        threadLabels,
        msg.isUnread(),
        msg.isStarred(),
        important,
        attachmentNames,
        `https://mail.google.com/mail/u/0/#inbox/${msg.getId()}`
      ]);
    });
  });

  // ----- 4. Append ------------------------------------------------------------
  if (data.length) {
    // find the last used row in column A only (ignores notes farther right)
    const lastDataRow = sheet.getRange('A:A').getLastRow();
    const startRow    = Math.max(3, lastDataRow + 1);

    sheet.getRange(startRow, 1, data.length, headers.length).setValues(data);

    // ----- 5. Full-width sort -------------------------------------------------
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 2) {
      sheet.getRange(3, 1, lastRow - 2, lastCol)
           .sort({ column: 2, ascending: false });   // sort by Date
    }
  }

  // ----- 6. Timestamp in A1 ---------------------------------------------------
  sheet.getRange(1, 1).setValue(`Updated: ${Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss')}`);
}

/**
 * Gets (or creates) a sheet by name
 */
function getOrCreateSheet(spreadsheet, name) {
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

/**
 * Parses an RFC-5322 address list into {names, emails}
 */
function parseAddressList(list) {
  if (!list) return { names: '', emails: '' };

  const parts  = list.split(',');
  const names  = [];
  const emails = [];

  parts.forEach(p => {
    const trimmed = p.trim();
    const m       = trimmed.match(/<(.*?)>/);

    if (m) {
      const email = m[1].trim();
      const name  = trimmed.replace(m[0], '').trim().replace(/^"|"$/g, '');
      names.push(name || email);
      emails.push(email);
    } else {
      names.push(trimmed);
      emails.push(trimmed);
    }
  });

  return { names: names.join('; '), emails: emails.join('; ') };
}
