/**
 * DataLayer.gs
 * The engine that connects your Google Apps Script logic to your Sheet Data.
 */

const DataLayer = {

  /**
   * Fetches all rows from a specified sheet key and returns them as an array of JSON objects.
   * @param {string} sheetKey - e.g., 'PROSPECTS', 'OUTREACH'
   * @returns {Array<Object>}
   */
  getRecords: function(sheetKey) {
    const sheetName = CONFIG.SHEETS[sheetKey];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) throw new Error(`Sheet ${sheetName} not found.`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // Empty sheet (only headers)

    const actualHeaders = data[0];
    const records = [];

    for (let i = 1; i < data.length; i++) {
      let record = { _rowNumber: i + 1 };
      let hasData = false;

      for (let j = 0; j < actualHeaders.length; j++) {
        let value = data[i][j];
        record[actualHeaders[j]] = value;
        if (value !== '') hasData = true;
      }

      if (hasData) records.push(record);
    }

    return records;
  },

  /**
   * OPTIMIZED: Fetches multiple sheets in a single spreadsheet connection.
   * Eliminates the N+1 read problem by opening the spreadsheet once.
   * @param {Array<string>} sheetKeys - Array of sheet keys e.g., ['PROSPECTS', 'OUTREACH']
   * @returns {Object} - Object with sheet keys as properties containing arrays of records
   */
  getMultipleRecords: function(sheetKeys) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = {};

    sheetKeys.forEach(sheetKey => {
      const sheetName = CONFIG.SHEETS[sheetKey];
      const sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        Logger.log(`Warning: Sheet ${sheetName} not found.`);
        result[sheetKey] = [];
        return;
      }

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        result[sheetKey] = [];
        return;
      }

      const actualHeaders = data[0];
      const records = [];

      for (let i = 1; i < data.length; i++) {
        let record = { _rowNumber: i + 1 };
        let hasData = false;

        for (let j = 0; j < actualHeaders.length; j++) {
          let value = data[i][j];
          record[actualHeaders[j]] = value;
          if (value !== '') hasData = true;
        }

        if (hasData) records.push(record);
      }

      result[sheetKey] = records;
    });

    return result;
  },

  /**
   * Saves a new record to the bottom of the requested sheet.
   * CONCURRENCY PROTECTED: Uses LockService to prevent data overwrites.
   * @param {string} sheetKey - e.g., 'OUTREACH'
   * @param {Object} recordData - JSON object with keys matching headers
   */
  addRecord: function(sheetKey, recordData) {
    const lock = LockService.getScriptLock();
    const lockTimeout = 30000; // 30 seconds

    try {
      // Acquire lock with timeout
      if (!lock.tryLock(lockTimeout)) {
        throw new Error('Could not acquire lock. Another process is writing data. Please try again.');
      }

      const sheetName = CONFIG.SHEETS[sheetKey];
      const headers = CONFIG.HEADERS[sheetKey];
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

      if (!sheet) throw new Error(`Sheet ${sheetName} not found.`);

      const rowToInsert = [];
      headers.forEach(header => {
        rowToInsert.push(recordData[header] !== undefined ? recordData[header] : "");
      });

      sheet.appendRow(rowToInsert);
      return { success: true, message: `Record added to ${sheetName}` };

    } catch (error) {
      Logger.log(`Error in addRecord: ${error.message}`);
      throw error;
    } finally {
      // Always release the lock
      lock.releaseLock();
    }
  }
};
