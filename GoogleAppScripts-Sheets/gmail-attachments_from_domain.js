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

function listAttachmentsFromDomain() {
  var domain = '@xyz.com/'; // Replace with the specific domain
  var threads = GmailApp.search('from:*' + domain); // Search for emails from the specified domain

  var sheetId = 'xyz'; // Replace with the actual Google Sheet ID
  var spreadsheet = SpreadsheetApp.openById(sheetId); // Open the spreadsheet by ID
  var sheetName = 'Email Attachments'; // Replace with the name of the sheet you want to use
  var sheet = spreadsheet.getSheetByName(sheetName); // Get the specific sheet by name

  // If the sheet doesn't exist, create it
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    // Add headers
    sheet.appendRow(['File Name', 'Email Subject', 'Date']);
  }

  // Iterate through the email threads
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var attachments = messages[j].getAttachments();
      for (var k = 0; k < attachments.length; k++) {
        // Add a new row for each attachment with its name, email subject, and date
        sheet.appendRow([attachments[k].getName(), messages[j].getSubject(), messages[j].getDate()]);
      }
    }
  }

  Logger.log('All attachments from emails sent from the domain ' + domain + ' have been listed in the specific Google Sheet.');
}
