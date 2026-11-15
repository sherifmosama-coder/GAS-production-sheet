function savePlan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plan = ss.getSheetById('1091805057');
  const oldPlans = ss.getSheetById('352632715');

  // 1. القراءة المجمعة (Batch Read)
  // تجميع القراءات المتفرقة لتقليل الاتصال بالسيرفر
  
  // قراءة المنطقة الأولى من شيت plan (من CM1 إلى CW2)
  // هذه المنطقة تحتوي على معظم المتغيرات والإعدادات
  const configData = plan.getRange('CM1:CW2').getValues();
  
  // قراءة المنطقة الثانية من شيت plan (A1:A3)
  const refData = plan.getRange('A1:A3').getValues();

  // قراءة الإحداثيات من شيت oldPlans (D1:E1)
  const targetData = oldPlans.getRange('D1:E1').getValues();

  // 2. تعيين المتغيرات من المصفوفات (Mapping)
  // Row 1 is index [0], Row 2 is index [1]
  // CM=0, CO=2, CQ=4, CR=5, CS=6, CT=7, CU=8, CW=10
  
  const chksmallErrorsMSG = configData[0][0]; // CM1
  const chksmallErrors = configData[1][0];    // CM2
  
  const currentPData = configData[1][2];      // CO2
  
  const pStRow = configData[1][4];            // CQ2
  const pStCol = configData[1][5];            // CR2
  const pNRows = configData[1][6];            // CS2
  const pNCols = configData[1][7];            // CT2
  
  const chkError = configData[1][8];          // CU2
  const chkSaved = configData[1][10];         // CW2 (موجود في الكود الأصلي لكن غير مستخدم في المنطق)

  // تعيين متغيرات المرجع (A1, A3)
  const refCol = refData[0][0]; // A1
  const refRow = refData[2][0]; // A3

  // تعيين متغيرات الهدف (D1, E1)
  const trgtRow = targetData[0][0]; // D1
  const trgtCol = targetData[0][1]; // E1

  // 3. التحقق والتنفيذ (Logic Execution)
  // تم الحفاظ على المنطق كما هو تماماً
  
  if (chkError == true && chksmallErrors == true) {
    // التحسين: لا نقوم بقراءة البيانات الكبيرة إلا إذا نجح التحقق
    const copyRangeValues = plan.getRange(pStRow, pStCol, pNRows, pNCols).getValues();
    const pasteRange = oldPlans.getRange(trgtRow, trgtCol, pNRows, pNCols);
    const refCell = oldPlans.getRange(refRow, refCol);

    pasteRange.setValues(copyRangeValues);
    refCell.setValue(currentPData);

    SpreadsheetApp.getActive().toast("تم حفظ بيانات الخطة بنجاح", "✅ SAVED !", 3);
  } 
  else {
    // معالجة الأخطاء (نفس المنطق الأصلي)
    // chkError usually holds 'true' or an 'Error Message String' based on the formula logic
    
    if (chkError != true && chksmallErrors != true) {
      SpreadsheetApp.getActive().toast(chkError + " " + chksmallErrorsMSG, "⚠️ NOT Saved !", 7);
    } 
    else if (chkError != true && chksmallErrors == true) {
      SpreadsheetApp.getActive().toast(chkError, "⚠️ NOT Saved !", 7);
    } 
    else {
      // الحالة المتبقية: chkError == true && chksmallErrors != true
      SpreadsheetApp.getActive().toast(chksmallErrorsMSG, "⚠️ NOT Saved !!", 7);
    }
  }
}
