// [UPDATED] closeDay - Instead of running logic, it just opens the Modal
function closeDay() {
  const html = HtmlService.createTemplateFromFile('CloseDayModal')
      .evaluate()
      .setWidth(450)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, ' '); // Empty title for clean look
}

// STEP 1: Validate Conditions (Returns true/false + message)
function step1_validate() {
  const sh = SpreadsheetApp.getActiveSpreadsheet();
  const ssh = sh.getSheetById("1626024145"); // Production
  
  // Read config Q3:Z6
  const configData = ssh.getRange("Q3:Z6").getValues();
  
  // Indices: Row 5 is index 2. Cols Q=0, R=1, S=2
  const misData_M = configData[2][0]; 
  const unsavedPallets_M = configData[2][1]; 
  const unsavedPlan_M = configData[2][2]; 
  const msg = configData[3][4]; // U6 (Row 6, Col U)

  if (misData_M === true && unsavedPallets_M === true && unsavedPlan_M === true) {
    return { success: true, msg: msg };
  } else {
    return { success: false, msg: msg };
  }
}

// STEP 2: Copy Report Data & Sign
function step2_copyReport() {
  const sh = SpreadsheetApp.getActiveSpreadsheet();
  const ssh = sh.getSheetById("1626024145"); // Production
  const tsh = sh.getSheetById("810212679");  // Report
  
  // Read config again for params
  const configData = ssh.getRange("Q3:Z6").getValues();

  const copyRow1 = configData[0][6]; // W3
  const copyCol1 = configData[0][7]; // X3
  const numRows1 = configData[0][8]; // Y3
  const numCols1 = configData[0][9]; // Z3
  const pasteRow1 = configData[1][6]; // W4
  const pasteCol1 = configData[1][7]; // X4
  const timeRow = configData[2][6]; // W5
  const timetCol = configData[2][7]; // X5

  // Execute Copy
  const reportRvalues = ssh.getRange(copyRow1, copyCol1, numRows1, numCols1).getValues();
  tsh.getRange(pasteRow1, pasteCol1, numRows1, numCols1).setValues(reportRvalues);

  // Sign Closure
  const userEmail = Session.getActiveUser().getEmail();
  const now = new Date();
  tsh.getRange(timeRow, timetCol, 1, 2).setValues([[userEmail, now]]);
  
  return "Done";
}

// STEP 3: Clear Data (Production & Plan)
function step3_clearData() {
  const sh = SpreadsheetApp.getActiveSpreadsheet();
  const ssh = sh.getSheetById("1626024145"); // Production
  const plan = sh.getSheetById("1091805057"); // Production Plan

  const planClearRanges = [
    'P4:P', 'R4:R', 'V4:V', 'Z4:Z', 'AD4:AD', 'AH4:AH', 'AL4:AL', 
    'AP4:AP', 'AT4:AT', 'AX4:AX', 'BB4:BB', 'BF4:BF', 'C4:F', 'H4:J'
  ];
  plan.getRangeList(planClearRanges).clearContent();

  const prodClearR = [
    'DU13:DV', 'DX13:ED', 'EJ13:EL'
  ];
  ssh.getRangeList(prodClearR).clearContent();
  
  return "Done";
}

// STEP 4: Update Consumption (Calls your optimized function)
function step4_updateConsumption() {
  // We call the function you already have. 
  // Ensure 'consumptionToMat' is defined in your script files.
  try {
    consumptionToMat(); 
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
