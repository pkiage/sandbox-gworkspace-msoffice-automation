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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu.
  ui.createMenu('Custom Menu')
      .addItem('Extract Links from Folder', 'showPrompt')
      .addToUi();
}

function showPrompt() {
  var ui = SpreadsheetApp.getUi();
  
  // First prompt to get the Google Drive folder link
  var folderResponse = ui.prompt('Enter Google Drive Folder Link', 
                           'Please enter the full link to the Google Drive folder:', 
                           ui.ButtonSet.OK_CANCEL);

  // Process the folder link response.
  if (folderResponse.getSelectedButton() == ui.Button.OK) {
    var url = folderResponse.getResponseText();
    var folderId = extractFolderId(url);
    
    // Second prompt to get the sheet name
    var sheetResponse = ui.prompt('Enter Sheet Name', 
                           'Please enter the name of the sheet where you want the file links:', 
                           ui.ButtonSet.OK_CANCEL);
    
    // Process the sheet name response.
    if (sheetResponse.getSelectedButton() == ui.Button.OK) {
      var sheetName = sheetResponse.getResponseText();
      if (folderId) {
        listFilesInFolder(folderId, sheetName);
      } else {
        ui.alert('Invalid Google Drive folder link. Please try again.');
      }
    }
  }
}

function extractFolderId(url) {
  var regex = /[-\w]{25,}/;
  var match = url.match(regex);
  return match ? match[0] : null;
}

function listFilesInFolder(folderId, sheetName) {
  var folder = DriveApp.getFolderById(folderId);
  
  // Get the specific sheet by name (using the input from the user)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    // If the sheet doesn't exist, create it.
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  } else {
    // Clear the sheet if it already exists.
    sheet.clear();
  }

  sheet.appendRow(["File Name", "URL", "Folder Path"]);
  
  // Call the recursive function to list all files and subfolders
  listAllFilesRecursive(folder, sheet, folder.getName());
}

function listAllFilesRecursive(folder, sheet, folderPath) {
  // Get all files in the current folder
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    sheet.appendRow([file.getName(), file.getUrl(), folderPath]);
  }
  
  // Recursively get all subfolders and their files
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var subfolderPath = folderPath + "/" + subfolder.getName();
    listAllFilesRecursive(subfolder, sheet, subfolderPath);
  }
}
