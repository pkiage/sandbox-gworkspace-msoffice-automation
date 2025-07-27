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
  ui.createMenu('Custom Menu')
    .addItem('Run Script', 'splitAndMove')
    .addToUi();
}

function splitAndMove() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('xyz'); // replace with your source sheet name
  var outputSheet = ss.getSheetByName('xyz'); // replace with your output sheet name
  
  // Create a temporary sheet and copy the visible cells there
  var tempSheet = ss.insertSheet('temp');
  sourceSheet.getDataRange().copyTo(tempSheet.getRange(1, 1), {contentsOnly: true});
  
  // Get the headers (first row) of the temp sheet
  var headers = tempSheet.getRange(1, 1, 1, tempSheet.getLastColumn()).getValues()[0];
  
  // Find the column index of 'Column name'
  var sourceColumnIndex = headers.indexOf('xyz'); // replace with your column name
  
  // If 'Investors' column is not found, log an error and exit the function
  if (sourceColumnIndex === -1) {
    Logger.log('Error: Column "xyz" not found.'); // replace "xyz" with your column name
    ss.deleteSheet(tempSheet); // Delete the temporary sheet before exiting
    return;
  }
  
  // Add 1 to the column index because SpreadsheetApp is 1-indexed
  sourceColumnIndex += 1;
  
  // Loop over the rows in the source data
  for (var i = 2; i <= tempSheet.getLastRow(); i++) {
    var cellData = tempSheet.getRange(i, sourceColumnIndex).getValue();
    
    // Split the cell data by comma
    var splitData = cellData.split(',');
    
    // Loop over the split data and output to the output sheet
    for (var j = 0; j < splitData.length; j++) {
      // Output each piece of split data to a new row in the output sheet
      outputSheet.appendRow([splitData[j].trim()]);
    }
  }
  
  // Delete the temporary sheet
  ss.deleteSheet(tempSheet);
}

