/**
 * DataLayer.gs
 * The engine that connects your Google Apps Script logic to your Sheet Data.
 * UPDATED: Now uses SchemaRegistry for fuzzy-matched column lookups
 */

const DataLayer = {

  /**
   * Fetches all rows from a specified sheet key and returns them as an array of JSON objects.
   * Uses SchemaRegistry for fuzzy-matched column headers.
   * @param {string} sheetKey - e.g., 'PROSPECTS', 'OUTREACH'
   * @returns {Array<Object>}
   */
  getRecords: function(sheetKey) {
    // Check if schema exists in SchemaRegistry
    const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
    
    if (hasSchema) {
      // Use smart schema-aware loading with fuzzy matching
      return this._getRecordsWithSchema(sheetKey);
    } else {
      // Fallback to legacy loading for sheets without schemas
      return this._getRecordsLegacy(sheetKey);
    }
  },

  /**
   * Schema-aware record loading with fuzzy-matched headers
   * @private
   */
  _getRecordsWithSchema: function(sheetKey) {
    const schema = SchemaRegistry.SCHEMA[sheetKey];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(schema.sheetName);

    if (!sheet) {
      Logger.log(`Warning: Sheet ${schema.sheetName} not found.`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const actualHeaders = data[0];
    const columnMap = SchemaRegistry.buildColumnMap(sheetKey, actualHeaders);
    
    const records = [];

    for (let i = 1; i < data.length; i++) {
      let record = { _rowNumber: i + 1 };
      let hasData = false;

      // Only map columns that are in our schema
      for (const [canonicalHeader, colIndex] of Object.entries(columnMap)) {
        let value = data[i][colIndex];
        record[canonicalHeader] = value;
        if (value !== '' && value !== null && value !== undefined) {
          hasData = true;
        }
      }

      if (hasData) records.push(record);
    }

    return records;
  },

  /**
   * Legacy record loading (for sheets without schema definitions)
   * @private
   */
  _getRecordsLegacy: function(sheetKey) {
    const sheetName = CONFIG.SHEETS[sheetKey];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`Warning: Sheet ${sheetName} not found.`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

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
   * Uses SchemaRegistry for fuzzy-matched column headers.
   * @param {Array<string>} sheetKeys - Array of sheet keys e.g., ['PROSPECTS', 'OUTREACH']
   * @returns {Object} - Object with sheet keys as properties containing arrays of records
   */
  getMultipleRecords: function(sheetKeys) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = {};

    sheetKeys.forEach(sheetKey => {
      try {
        const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
        
        if (hasSchema) {
          // Schema-aware loading
          const schema = SchemaRegistry.SCHEMA[sheetKey];
          const sheet = ss.getSheetByName(schema.sheetName);

          if (!sheet) {
            Logger.log(`Warning: Sheet ${schema.sheetName} not found.`);
            result[sheetKey] = [];
            return;
          }

          const data = sheet.getDataRange().getValues();
          if (data.length < 2) {
            result[sheetKey] = [];
            return;
          }

          const actualHeaders = data[0];
          const columnMap = SchemaRegistry.buildColumnMap(sheetKey, actualHeaders);
          const records = [];

          for (let i = 1; i < data.length; i++) {
            let record = { _rowNumber: i + 1 };
            let hasData = false;

            for (const [canonicalHeader, colIndex] of Object.entries(columnMap)) {
              let value = data[i][colIndex];
              record[canonicalHeader] = value;
              if (value !== '' && value !== null && value !== undefined) {
                hasData = true;
              }
            }

            if (hasData) records.push(record);
          }

          result[sheetKey] = records;
        } else {
          // Legacy loading
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
        }
      } catch (e) {
        Logger.log(`Error loading ${sheetKey}: ${e.message}`);
        result[sheetKey] = [];
      }
    });

    return result;
  },

  /**
   * Saves a new record to the bottom of the requested sheet.
   * CONCURRENCY PROTECTED: Uses LockService to prevent data overwrites.
   * Uses SchemaRegistry to map canonical field names to actual column positions.
   * @param {string} sheetKey - e.g., 'OUTREACH'
   * @param {Object} recordData - JSON object with canonical field names
   */
  addRecord: function(sheetKey, recordData) {
    const lock = LockService.getScriptLock();
    const lockTimeout = 30000; // 30 seconds

    try {
      // Acquire lock with timeout
      if (!lock.tryLock(lockTimeout)) {
        throw new Error('Could not acquire lock. Another process is writing data. Please try again.');
      }

      const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
      
      if (hasSchema) {
        // Schema-aware insertion
        const schema = SchemaRegistry.SCHEMA[sheetKey];
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(schema.sheetName);

        if (!sheet) throw new Error(`Sheet ${schema.sheetName} not found.`);

        // Get actual headers from sheet
        const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const columnMap = SchemaRegistry.buildColumnMap(sheetKey, actualHeaders);

        // Build row with correct column positions
        const maxCol = Math.max(...Object.values(columnMap)) + 1;
        const rowToInsert = new Array(maxCol).fill('');

        for (const [fieldName, value] of Object.entries(recordData)) {
          const colIndex = columnMap[fieldName];
          if (colIndex !== undefined) {
            rowToInsert[colIndex] = value;
          } else {
            // Try fuzzy match for fields not in schema
            const matchedField = SchemaRegistry.getCanonicalHeader(sheetKey, fieldName);
            if (matchedField && columnMap[matchedField] !== undefined) {
              rowToInsert[columnMap[matchedField]] = value;
            }
          }
        }

        sheet.appendRow(rowToInsert);
        return { success: true, message: `Record added to ${schema.sheetName}` };
      } else {
        // Legacy insertion using CONFIG.HEADERS
        const sheetName = CONFIG.SHEETS[sheetKey];
        const headers = CONFIG.HEADERS[sheetKey];
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

        if (!sheet) throw new Error(`Sheet ${sheetName} not found.`);

        const rowToInsert = [];
        headers.forEach(header => {
          // Try exact match first, then fuzzy match
          let value = recordData[header];
          if (value === undefined) {
            // Try to find a fuzzy match in recordData keys
            for (const [key, val] of Object.entries(recordData)) {
              if (SchemaRegistry.fuzzyMatchHeader(key, [header])) {
                value = val;
                break;
              }
            }
          }
          rowToInsert.push(value !== undefined ? value : '');
        });

        sheet.appendRow(rowToInsert);
        return { success: true, message: `Record added to ${sheetName}` };
      }

    } catch (error) {
      Logger.log(`Error in addRecord: ${error.message}`);
      throw error;
    } finally {
      // Always release the lock
      lock.releaseLock();
    }
  },

  /**
   * Update an existing record by row number
   * Uses SchemaRegistry for fuzzy-matched column positioning
   * @param {string} sheetKey - e.g., 'PROSPECTS'
   * @param {number} rowNumber - Row number to update
   * @param {Object} updates - Object with canonical field names and new values
   * @returns {Object} - Result of update
   */
  updateRecord: function(sheetKey, rowNumber, updates) {
    const lock = LockService.getScriptLock();
    
    try {
      if (!lock.tryLock(30000)) {
        throw new Error('Could not acquire lock');
      }

      const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
      
      if (hasSchema) {
        const schema = SchemaRegistry.SCHEMA[sheetKey];
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(schema.sheetName);
        
        if (!sheet) throw new Error(`Sheet ${schema.sheetName} not found`);

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const columnMap = SchemaRegistry.buildColumnMap(sheetKey, headers);

        // Update each field
        for (const [fieldName, value] of Object.entries(updates)) {
          const colIndex = columnMap[fieldName];
          if (colIndex !== undefined) {
            sheet.getRange(rowNumber, colIndex + 1).setValue(value);
          } else {
            // Try fuzzy match
            const matchedField = SchemaRegistry.getCanonicalHeader(sheetKey, fieldName);
            if (matchedField && columnMap[matchedField] !== undefined) {
              sheet.getRange(rowNumber, columnMap[matchedField] + 1).setValue(value);
            }
          }
        }
      } else {
        // Legacy update
        const sheetName = CONFIG.SHEETS[sheetKey];
        const headers = CONFIG.HEADERS[sheetKey];
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        
        if (!sheet) throw new Error(`Sheet ${sheetName} not found`);

        for (const [fieldName, value] of Object.entries(updates)) {
          const colIndex = headers.indexOf(fieldName);
          if (colIndex >= 0) {
            sheet.getRange(rowNumber, colIndex + 1).setValue(value);
          } else {
            // Try fuzzy match against headers
            const matched = SchemaRegistry.fuzzyMatchHeader(fieldName, headers);
            if (matched) {
              const matchedIndex = headers.indexOf(matched);
              sheet.getRange(rowNumber, matchedIndex + 1).setValue(value);
            }
          }
        }
      }

      return { success: true, message: `Row ${rowNumber} updated` };

    } catch (error) {
      Logger.log(`Error in updateRecord: ${error.message}`);
      throw error;
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Get column index for a field (with fuzzy matching)
   * @param {string} sheetKey - Schema key
   * @param {string} fieldName - Field name (can be approximate)
   * @returns {number} - Column index (-1 if not found)
   */
  getColumnIndex: function(sheetKey, fieldName) {
    const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
    
    if (hasSchema) {
      const schema = SchemaRegistry.SCHEMA[sheetKey];
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(schema.sheetName);
      if (!sheet) return -1;
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      return SchemaRegistry.getColumnIndex(sheetKey, headers, fieldName);
    } else {
      const sheetName = CONFIG.SHEETS[sheetKey];
      const headers = CONFIG.HEADERS[sheetKey];
      
      // Try exact match first
      let index = headers.indexOf(fieldName);
      if (index >= 0) return index;
      
      // Try fuzzy match
      const matched = SchemaRegistry.fuzzyMatchHeader(fieldName, headers);
      if (matched) {
        return headers.indexOf(matched);
      }
      
      return -1;
    }
  },

  /**
   * Get the value of a specific cell by row and field name
   * @param {string} sheetKey - Schema key
   * @param {number} rowNumber - Row number
   * @param {string} fieldName - Field name (fuzzy matched)
   * @returns {*} - Cell value
   */
  getCellValue: function(sheetKey, rowNumber, fieldName) {
    const colIndex = this.getColumnIndex(sheetKey, fieldName);
    if (colIndex < 0) return null;
    
    const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
    const sheetName = hasSchema ? SchemaRegistry.SCHEMA[sheetKey].sheetName : CONFIG.SHEETS[sheetKey];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) return null;
    return sheet.getRange(rowNumber, colIndex + 1).getValue();
  },

  /**
   * Set the value of a specific cell by row and field name
   * @param {string} sheetKey - Schema key
   * @param {number} rowNumber - Row number
   * @param {string} fieldName - Field name (fuzzy matched)
   * @param {*} value - Value to set
   */
  setCellValue: function(sheetKey, rowNumber, fieldName, value) {
    const colIndex = this.getColumnIndex(sheetKey, fieldName);
    if (colIndex < 0) {
      throw new Error(`Field '${fieldName}' not found in ${sheetKey}`);
    }
    
    const hasSchema = SchemaRegistry.SCHEMA[sheetKey];
    const sheetName = hasSchema ? SchemaRegistry.SCHEMA[sheetKey].sheetName : CONFIG.SHEETS[sheetKey];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} not found`);
    }
    
    sheet.getRange(rowNumber, colIndex + 1).setValue(value);
  }
};
