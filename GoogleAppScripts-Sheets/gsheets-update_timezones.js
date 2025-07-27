/*
 * Copyright 2025 The Contributors
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
Overview

This script allows you to set the timezone for all Google Spreadsheets within a specific folder and its subfolders. You can enter a full folder link (e.g., https://drive.google.com/drive/folders/xyz), and the script will automatically extract the folder ID and apply the timezone to each spreadsheet.

How to Use

After reloading, a new menu item called Custom Timezone Tools will appear in the menu bar (next to Help).

Click Custom Timezone Tools > Set Timezone for Sheets.

A prompt will appear asking for the Google Drive folder link: Enter the full link to the folder (e.g., https://drive.google.com/drive/folders/xyz).

Another prompt will ask for the timezone (e.g., Africa/Nairobi, America/New_York): Enter the timezone in the correct format.

The script will then change the timezone for al  
Troubleshooting

Invalid Folder Link: Make sure the folder link is in the correct format (e.g., https://drive.google.com/drive/folders/<folderID>).
Permissions Error: If you encounter a permissions error, ensure you have authorized the script and that it has access to Google Drive.
Timezone Format: Use the correct timezone format (e.g., Africa/Nairobi or America/New_York). You can find valid timezones here.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Timezone Tools')
    .addItem('Set Timezone for Sheets', 'promptForFolderLinkAndTimeZone')
    .addToUi();
}

function promptForFolderLinkAndTimeZone() {
  var ui = SpreadsheetApp.getUi();
  
  // Prompt the user for folder link
  var folderResponse = ui.prompt('Set Timezone for Sheets', 'Please enter the Folder Link:', ui.ButtonSet.OK_CANCEL);
  if (folderResponse.getSelectedButton() == ui.Button.OK) {
    var folderLink = folderResponse.getResponseText();
    var folderId = extractFolderIdFromLink(folderLink); // Extract folder ID from link
    
    if (folderId) {
      // Prompt the user for timezone
      var timeZoneResponse = ui.prompt('Set Timezone for Sheets', 'Please enter the Timezone (e.g., Africa/Nairobi):', ui.ButtonSet.OK_CANCEL);
      if (timeZoneResponse.getSelectedButton() == ui.Button.OK) {
        var timeZone = timeZoneResponse.getResponseText();
        
        // Run the function to set the timezone for all spreadsheets in the folder and its subfolders
        setTimeZoneForAllSheetsInFolderAndSubfolders(folderId, timeZone);
        
        ui.alert('Timezone updated successfully!');
      }
    } else {
      ui.alert('Invalid folder link. Please make sure the link is correct.');
    }
  }
}

function extractFolderIdFromLink(link) {
  // Extract folder ID using regex (folder ID is the part after 'folders/')
  var regex = /\/folders\/([a-zA-Z0-9-_]+)/;
  var match = link.match(regex);
  if (match && match[1]) {
    return match[1];
  } else {
    return null;
  }
}

function setTimeZoneForAllSheetsInFolderAndSubfolders(folderId, timeZone) {
  try {
    var folder = DriveApp.getFolderById(folderId); // Correct method to get folder by ID
    processFolder(folder, timeZone); // Process folder
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    SpreadsheetApp.getUi().alert('An error occurred: ' + e.message);
  }
}

function processFolder(folder, timeZone) {
  // Get all files in the current folder
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  while (files.hasNext()) {
    var file = files.next();
    var spreadsheet = SpreadsheetApp.open(file);
    
    // Set the timezone
    spreadsheet.setSpreadsheetTimeZone(timeZone);
    Logger.log('Updated: ' + file.getName() + ' to timezone ' + timeZone);
  }
  
  // Get all subfolders in the current folder and process them recursively
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    processFolder(subfolder, timeZone);
  }
}
