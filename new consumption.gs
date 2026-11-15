// Entry trigger to launch the interactive modal dialog
function consumptionToMat() {
  const html = HtmlService.createHtmlOutputFromFile('ConsumptionModal')
      .setWidth(520)
      .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sync Material Consumption');
}

// 1. Lightweight Scanner: Scans Rows 2 & 3 without loading massive data
function getConsumptionRanges() {
  const SSS_ID = '1J7HG2t5EGc9fCmzAi6Nz2xbaaTarMHvbB3_PCFYNY0I';
  const sourceSS = SpreadsheetApp.openById(SSS_ID);
  const srcSheet = sourceSS.getSheetById(1289381484);

  if (!srcSheet) return { error: 'Source sheet not found.' };

  const lastCol = srcSheet.getLastColumn();
  const maxCols = lastCol - 3; // Columns from D onward

  if (maxCols < 1) return { ranges: [] };

  // Read only Row 2 (Flags) and Row 3 (Dates) across all columns
  const headerData = srcSheet.getRange(2, 4, 2, maxCols).getValues();
  const flags = headerData[0];
  const dates = headerData[1];

  const blocks = [];
  let currentBlock = null;

  const formatDate = (val) => {
    if (!val) return 'Unknown Date';
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone() || "GMT", "dd MMM yyyy");
    }
    return String(val);
  };

  for (let i = 0; i < flags.length; i++) {
    if (flags[i] === true) {
      let dateStr = formatDate(dates[i]);
      if (!currentBlock) {
        currentBlock = {
          startCol: i + 4, // 1-based index (Col D = 4)
          numCols: 1,
          startDate: dateStr,
          endDate: dateStr
        };
      } else {
        currentBlock.numCols++;
        currentBlock.endDate = dateStr;
      }
    } else {
      if (currentBlock) {
        blocks.push(currentBlock);
        currentBlock = null;
      }
    }
  }
  if (currentBlock) blocks.push(currentBlock);

  // Build clean display labels for the UI
  const ranges = blocks.map((b, idx) => {
    let label = b.startDate === b.endDate 
      ? `${b.startDate} (${b.numCols} Day)`
      : `${b.startDate} → ${b.endDate} (${b.numCols} Days)`;
    return {
      id: 'range_' + idx,
      startCol: b.startCol,
      numCols: b.numCols,
      label: label
    };
  });

  return { ranges: ranges };
}

// 2. Targeted Batch Writer: Copies only one specific chunk
function processSingleConsumptionBatch(startCol, numCols) {
  const SSS_ID = '1J7HG2t5EGc9fCmzAi6Nz2xbaaTarMHvbB3_PCFYNY0I';
  const TSS_ID = '1AATh8aLorXF428tbExFU6LjWAdsn8ybQjp4933BeTzU';

  const srcSheet = SpreadsheetApp.openById(SSS_ID).getSheetById(1289381484);
  const tgtSheet = SpreadsheetApp.openById(TSS_ID).getSheetById(740404196);

  const lastRow = srcSheet.getLastRow();
  if (lastRow < 5) return { success: false, message: 'No rows to copy.' };

  const numRows = lastRow - 4;

  // A. Fetch only the targeted range slice
  const sliceData = srcSheet.getRange(5, startCol, numRows, numCols).getValues();

  // B. Write payload slice to target
  tgtSheet.getRange(5, startCol, numRows, numCols).setValues(sliceData);

  // C. Write metadata headers in single-row bulk arrays
  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date();

  tgtSheet.getRange(2, startCol, 1, numCols).setValues([new Array(numCols).fill(userEmail)]);
  tgtSheet.getRange(4, startCol, 1, numCols).setValues([new Array(numCols).fill(timestamp)]);

  return { success: true };
}
