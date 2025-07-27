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
Columns to add
- Master name (script ok?) Yes/No
- [if applicable] why change (name adjustment/ grouping adjustment)
- Human note
- Adjusted name
- Final name =IF(G7="Yes",D7,J7)
G7 is if script name ok whereas D7 output name and J7 is adjusted name
*/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Fuzzy matching")
    .addItem("Group similar names", "startFuzzyGrouping")
    .addToUi();
}

function startFuzzyGrouping() {
  const ui = SpreadsheetApp.getUi();

  const colPrompt = ui.prompt(
    "Fuzzy Grouping",
    "Which column contains the enterprise names? (e.g., B)",
    ui.ButtonSet.OK_CANCEL
  );
  if (colPrompt.getSelectedButton() !== ui.Button.OK) return;
  const columnLetter = colPrompt.getResponseText().toUpperCase().trim();

  const thresholdPrompt = ui.prompt(
    "Similarity Threshold (0–100)",
    "Enter a number (recommended: 85 = strict, 70 = lenient, 60 = very loose):",
    ui.ButtonSet.OK_CANCEL
  );
  if (thresholdPrompt.getSelectedButton() !== ui.Button.OK) return;

  let threshold = parseInt(thresholdPrompt.getResponseText().trim(), 10);
  if (isNaN(threshold) || threshold < 0 || threshold > 100) threshold = 70; // fallback default

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const header = data[0];
  const nameColIndex = columnLetter.charCodeAt(0) - 65;
  const idColIndex = header.findIndex(h => h.toString().toLowerCase().includes("id"));

  if (nameColIndex < 0 || nameColIndex >= header.length) {
    ui.alert("Invalid column letter.");
    return;
  }

  const rawNames = data.slice(1).map(row => row[nameColIndex]);
  const ids = idColIndex !== -1 ? data.slice(1).map(row => row[idColIndex]) : Array(rawNames.length).fill("");

  function cleanName(name) {
    return name.toLowerCase()
      .replace(/\b(ltd|limited|plc|inc|corp|co|company|enterprises?)\b/g, '')
      .replace(/[^\w\s]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function tokenSetRatio(a, b) {
    const aTokens = new Set(cleanName(a).split(/\s+/));
    const bTokens = new Set(cleanName(b).split(/\s+/));
    const shared = new Set([...aTokens].filter(x => bTokens.has(x)));

    const intersection = [...shared].join(' ');
    const aJoin = [...aTokens].join(' ');
    const bJoin = [...bTokens].join(' ');

    const base = intersection.length;
    const maxLen = Math.max(aJoin.length, bJoin.length);
    let ratio = (base / maxLen) * 100;

    // Fallback: if cleaned no-space versions match, force 100%
    const aNoSpace = cleanName(a).replace(/\s+/g, '');
    const bNoSpace = cleanName(b).replace(/\s+/g, '');
    if (aNoSpace === bNoSpace) ratio = Math.max(ratio, 100);

    return ratio;
  }

  const cleanedNames = rawNames.map(cleanName);
  const groups = [];
  const assigned = Array(rawNames.length).fill(false);

  for (let i = 0; i < cleanedNames.length; i++) {
    if (assigned[i]) continue;
    let group = [i];
    assigned[i] = true;
    for (let j = i + 1; j < cleanedNames.length; j++) {
      if (!assigned[j] && tokenSetRatio(rawNames[i], rawNames[j]) >= threshold) {
        group.push(j);
        assigned[j] = true;
      }
    }
    groups.push(group);
  }

  const output = [["Group ID", "Enterprise Name", "Cleaned Name", "Master Name", "ID", "Similarity (%)"]];

  groups.forEach((group, idx) => {
    const groupOriginals = group.map(i => rawNames[i]);
    const groupCleaned = group.map(i => cleanedNames[i]);
    const groupIDs = group.map(i => ids[i]);

    const masterIndex = groupCleaned
      .map((val, i) => [val.length, i])
      .sort((a, b) => a[0] - b[0])[0][1];

    const masterName = groupOriginals[masterIndex];

    group.forEach(i => {
      const score = tokenSetRatio(rawNames[i], masterName);
      output.push([
        "Group " + (idx + 1),
        rawNames[i],
        cleanedNames[i],
        masterName,
        ids[i],
        score.toFixed(1)
      ]);
    });
  });

  let outputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Grouped Matches");
  if (!outputSheet) {
    outputSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Grouped Matches");
  } else {
    outputSheet.clear();
  }

  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}
