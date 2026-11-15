function savePallet() {
  const sh = SpreadsheetApp.openById('1J7HG2t5EGc9fCmzAi6Nz2xbaaTarMHvbB3_PCFYNY0I');
  const ssh = sh.getSheetById("1626024145"); // Source Sheet (Production)
  const tsh = sh.getSheetById("415218049");  // Target Sheet (Pallets Out)
  const user = Session.getActiveUser().getEmail(); 

  // 1. BATCH READ CONFIGURATION
  // Instead of calling .getValue() 10 times, we get all config data (P8:S11) at once.
  // Mapping Indices based on range P8:S11:
  // Row 8 = index 0 (Target Coords)
  // Row 9 = index 1 (First Table Coords)
  // Row 10 = index 2 (Empty/Ignored)
  // Row 11 = index 3 (Full Table Coords)
  // Columns: P=0, Q=1, R=2, S=3
  
  const configData = ssh.getRange("P8:S11").getValues();

  // Extract First Table Parameters (Row 9 -> index 1)
  const numRows1 = configData[1][2]; // R9
  
  // Check if number of rows is zero immediately
  if (numRows1 === 0) { 
    sh.toast("No pallets ready to save", "⚠️ Unable to save", 3);
    return;
  }

  const startRow1 = configData[1][0]; // P9
  const startCol1 = configData[1][1]; // Q9
  const numCols1  = configData[1][3]; // S9

  // Extract Full Table Parameters (Row 11 -> index 3)
  const startRow2 = configData[3][0]; // P11
  const startCol2 = configData[3][1]; // Q11
  const numRows2  = configData[3][2]; // R11
  const numCols2  = configData[3][3]; // S11

  // Extract Target Parameters (Row 8 -> index 0)
  const startRowT = configData[0][0]; // P8
  const startColT = configData[0][1]; // Q8

  // 2. PROCESS TIMESTAMPS (First Table)
  const firstTableRange = ssh.getRange(startRow1, startCol1, numRows1, numCols1);
  // Timestamp range is 2 columns wide, starting immediately after the data columns
  const timestampUserRange = ssh.getRange(startRow1, startCol1 + numCols1, numRows1, 2);
  
  const data = firstTableRange.getValues();
  const timestampUserData = timestampUserRange.getValues();
  const now = new Date();

  // Loop to update data in memory
  let hasUpdates = false;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== "") { 
      timestampUserData[i][0] = now;  // Insert timestamp
      timestampUserData[i][1] = user; // Insert user
      hasUpdates = true;
    }
  }

  // Write timestamps back to Source Sheet
  if (hasUpdates) {
    timestampUserRange.setValues(timestampUserData);
  }

  // 3. COPY TO TARGET
  // We read the full range NOW to ensure we capture the timestamps we just wrote
  const fullRange = ssh.getRange(startRow2, startCol2, numRows2, numCols2);
  const dataToCopy = fullRange.getValues();
  
  const targetRange = tsh.getRange(startRowT, startColT, dataToCopy.length, dataToCopy[0].length);
  targetRange.setValues(dataToCopy);

  sh.toast(`${numRows1} pallets saved`, "✅ Save Successful", 3);
}
