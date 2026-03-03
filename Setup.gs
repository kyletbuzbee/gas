/**
 * Setup.gs
 * Run the `initializeFreshCRM` function ONCE to set up your Google Sheet tabs and headers.
 */
function initializeFreshCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Loop through all defined sheets in the config
  for (const key in CONFIG.SHEETS) {
    const sheetName = CONFIG.SHEETS[key];
    const headers = CONFIG.HEADERS[key]; // Fetch exact headers from Config
    
    let sheet = ss.getSheetByName(sheetName);
    
    // If sheet doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log(`Created new sheet: ${sheetName}`);
    }
    
    // If we have headers defined for this sheet, inject them
    if (headers && headers.length > 0) {
      // Clear existing headers just in case it's a reset, then set new ones
      sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      
      // Freeze the top row
      sheet.setFrozenRows(1);
    }
  }
  
  // Clean up default "Sheet1" if it's empty
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }
  
  SpreadsheetApp.getUi().alert('Success: Fresh CRM Architecture Initialized!');
}
