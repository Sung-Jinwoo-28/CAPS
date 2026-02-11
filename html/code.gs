function syncNonChristitesToData(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SOURCE_SHEET_NAME = "Non-Christites";
  const TARGET_SHEET_NAME = "Data";

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (!sourceSheet || !targetSheet) return;

  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  if (sourceData.length < 2) return;

  const sourceHeaders = sourceData[0];
  const targetHeaders = targetData[0];

  const sIdx = headerIndexMap(sourceHeaders);
  const tIdx = headerIndexMap(targetHeaders);

  // Build lookup map for Data sheet using Attendee ID
  const dataIdMap = {};
  for (let i = 1; i < targetData.length; i++) {
    const id = targetData[i][tIdx["ID"]];
    if (id) dataIdMap[id] = i;
  }

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const attendeeId = row[sIdx["Attendee ID"]];
    if (!attendeeId) continue;

    const mappedRow = Array(targetHeaders.length).fill("");

    mappedRow[tIdx["ID"]] = attendeeId;
    mappedRow[tIdx["Name"]] = row[sIdx["Name"]];
    mappedRow[tIdx["Contact"]] = row[sIdx["Email | Phone"]];
    mappedRow[tIdx["Category"]] = row[sIdx["College Name"]];
    mappedRow[tIdx["Workshop"]] = row[sIdx["Workshop"]];
    mappedRow[tIdx["QR Link"]] = row[sIdx["QR Code URL"]];
    mappedRow[tIdx["Checked In"]] = row[sIdx["Checked In"]];
    mappedRow[tIdx["Kit Given"]] = row[sIdx["Kit Given"]];
    mappedRow[tIdx["Snack 11AM"]] = row[sIdx["Snack 11AM"]];
    mappedRow[tIdx["Snack 1PM"]] = row[sIdx["Snack 1PM"]];
    mappedRow[tIdx["Timestamp"]] = row[sIdx["Timestamp"]];
    mappedRow[tIdx["Payment Doc"]] = row[sIdx["Payment Proof"]];

    if (dataIdMap[attendeeId] !== undefined) {
      // Update existing row
      targetSheet
        .getRange(dataIdMap[attendeeId] + 1, 1, 1, mappedRow.length)
        .setValues([mappedRow]);
    } else {
      // Insert new row
      targetSheet.appendRow(mappedRow);
    }
  }
}

// Utility: create header → index map
function headerIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}
