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
 * Adds “Folders” menu on spreadsheet open
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Folders')
    .addItem('Create hierarchy', 'makeFolders')
    .addToUi();
}

/**
 * Main entry point – build folder tree under a user‑chosen parent
 */
function makeFolders() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Parent folder',
    'Paste the Google Drive folder URL or ID where you want the hierarchy created:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;                // user cancelled

  const parentId = extractId_(resp.getResponseText().trim());
  if (!parentId) { ui.alert('⚠️ Could not detect a valid folder ID.'); return; }

  let parent;
  try { parent = DriveApp.getFolderById(parentId); }
  catch (e) { ui.alert('⚠️ Drive says that folder ID is invalid or you lack access.'); return; }

  const sheet = SpreadsheetApp.getActiveSheet();
  const rows  = sheet.getDataRange().getValues().slice(1);              // skip header

  rows.forEach(([rootName, subStr]) => {
    if (!rootName) return;                                              // skip blanks
    const rootFolder = getOrCreate_(parent, rootName);

    if (subStr) {
      subStr.split(/[,/]/).map(s => s.trim()).forEach(levelPath => {
        if (levelPath) getOrCreate_(rootFolder, levelPath);
      });
    }
  });

  ui.alert('Folder hierarchy created under: ' + parent.getName());
}

/**
 * If user pasted a full URL, pull out the long ID; else return the string unchanged.
 */
function extractId_(input) {
  const match = input.match(/[-\w]{25,}/);   // Drive IDs are 25+ url‑safe chars
  return match ? match[0] : null;
}

/**
 * Returns existing folder “name” under “parent”, or creates it if absent
 */
function getOrCreate_(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}
