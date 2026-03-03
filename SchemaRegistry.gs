/**
 * SchemaRegistry.gs
 * Central schema management with fuzzy matching for column headers
 * All column lookups go through this to handle variations and prevent errors
 */

const SchemaRegistry = {
  // Canonical schema definitions - the single source of truth
  SCHEMA: {
    PROSPECTS: {
      sheetName: 'Prospects',
      required: ['Company ID', 'Company Name'],
      optional: ['Address', 'City', 'Zip Code', 'Industry', 'Latitude', 'Longitude', 
                 'Last Outcome', 'Last Outreach Date', 'Days Since Last Contact',
                 'Next Step Due Countdown', 'Next Steps Due Date', 'Contact Status',
                 'Close Probability', 'Priority Score', 'UrgencyBand', 'Urgency Score', 'Totals'],
      all: ['Company ID', 'Address', 'City', 'Zip Code', 'Company Name', 'Industry', 
            'Latitude', 'Longitude', 'Last Outcome', 'Last Outreach Date', 
            'Days Since Last Contact', 'Next Step Due Countdown', 'Next Steps Due Date', 
            'Contact Status', 'Close Probability', 'Priority Score', 
            'UrgencyBand', 'Urgency Score', 'Totals']
    },
    ACTIVE_CONTAINERS: {
      sheetName: 'Active',
      required: ['Company ID', 'Company Name'],
      optional: ['Location Name', 'Location Address', 'City', 'Zip Code', 
                 'Current Deployed Asset(s)', 'Container Size'],
      all: ['Company ID', 'Company Name', 'Location Name', 'Location Address', 
            'City', 'Zip Code', 'Current Deployed Asset(s)', 'Container Size']
    },
    OUTREACH: {
      sheetName: 'Outreach',
      required: ['Outreach ID', 'Company ID', 'Company'],
      optional: ['Visit Date', 'Notes', 'Outcome', 'Stage', 'Status', 
                 'Next Visit Date', 'Days Since Last Visit', 'Next Visit Countdown',
                 'Outcome Category', 'Follow Up Action', 'Owner', 'Prospects Match',
                 'Contact Type', 'Email Sent', 'Competitor'],
      all: ['Outreach ID', 'Company ID', 'Company', 'Visit Date', 'Notes', 
            'Outcome', 'Stage', 'Status', 'Next Visit Date', 'Days Since Last Visit',
            'Next Visit Countdown', 'Outcome Category', 'Follow Up Action', 
            'Owner', 'Prospects Match', 'Contact Type', 'Email Sent', 'Competitor']
    },
    SALES: {
      sheetName: 'Sales',
      required: ['Sales ID', 'Company ID', 'Company Name'],
      optional: ['Date', 'Material', 'Weight', 'Price', 'Payment Amount'],
      all: ['Sales ID', 'Date', 'Company Name', 'Company ID', 'Material', 
            'Weight', 'Price', 'Payment Amount']
    },
    ACCOUNTS: {
      sheetName: 'Accounts',
      required: ['Company Name'],
      optional: ['Deployed', 'Timestamp', 'Contact Name', 'Contact Phone', 
                 'Contact Role', 'Site Location', 'Mailing Location', 
                 'Roll-Off Fee', 'Handling of Metal', 'Roll Off Container Size', 
                 'Notes', 'Payout Price'],
      all: ['Deployed', 'Timestamp', 'Company Name', 'Contact Name', 'Contact Phone',
            'Contact Role', 'Site Location', 'Mailing Location', 'Roll-Off Fee',
            'Handling of Metal', 'Roll Off Container Size', 'Notes', 'Payout Price']
    },
    CONTACTS: {
      sheetName: 'Contacts',
      required: ['Name'],
      optional: ['Company', 'Account', 'Role', 'Department', 
                 'Phone Number', 'Email', 'Address'],
      all: ['Name', 'Company', 'Account', 'Role', 'Department', 
            'Phone Number', 'Email', 'Address']
    },
    TRANSACTIONS: {
      sheetName: 'Transactions',
      required: ['Transaction ID'],
      optional: ['Date', 'Company', 'Material', 'Net Weight', 'Price', 'Payment Amount'],
      all: ['Transaction ID', 'Date', 'Company', 'Material', 
            'Net Weight', 'Price', 'Payment Amount']
    }
  },

  // Cache for fuzzy-matched column indices
  _columnMapCache: {},

  /**
   * Fuzzy match a header name against canonical headers
   * @param {string} actualHeader - The header found in the sheet
   * @param {Array<string>} canonicalHeaders - The expected headers
   * @returns {string|null} - The matched canonical header or null
   */
  fuzzyMatchHeader: function(actualHeader, canonicalHeaders) {
    if (!actualHeader) return null;
    
    const normalized = this._normalize(actualHeader);
    
    // First try exact match
    for (const canonical of canonicalHeaders) {
      if (this._normalize(canonical) === normalized) {
        return canonical;
      }
    }
    
    // Try fuzzy match with similarity threshold
    let bestMatch = null;
    let bestScore = 0;
    
    for (const canonical of canonicalHeaders) {
      const score = this._similarity(normalized, this._normalize(canonical));
      if (score > bestScore && score >= 0.85) { // 85% similarity threshold
        bestScore = score;
        bestMatch = canonical;
      }
    }
    
    return bestMatch;
  },

  /**
   * Build a column map for a sheet - maps canonical headers to actual column indices
   * @param {string} sheetKey - The schema key (e.g., 'PROSPECTS')
   * @param {Array<string>} actualHeaders - Headers found in the sheet
   * @returns {Object} - Map of canonical headers to column indices
   */
  buildColumnMap: function(sheetKey, actualHeaders) {
    const cacheKey = sheetKey + '_' + actualHeaders.join(',');
    if (this._columnMapCache[cacheKey]) {
      return this._columnMapCache[cacheKey];
    }

    const schema = this.SCHEMA[sheetKey];
    if (!schema) {
      throw new Error(`Unknown schema: ${sheetKey}`);
    }

    const columnMap = {};
    const unmatched = [];
    
    actualHeaders.forEach((header, index) => {
      const matched = this.fuzzyMatchHeader(header, schema.all);
      if (matched) {
        columnMap[matched] = index;
      } else {
        unmatched.push(header);
      }
    });

    // Log warnings for unmatched headers
    if (unmatched.length > 0) {
      Logger.log(`Warning: Unmatched headers in ${sheetKey}: ${unmatched.join(', ')}`);
    }

    // Verify required fields are present
    const missing = schema.required.filter(req => !(req in columnMap));
    if (missing.length > 0) {
      Logger.log(`Critical: Missing required headers in ${sheetKey}: ${missing.join(', ')}`);
    }

    this._columnMapCache[cacheKey] = columnMap;
    return columnMap;
  },

  /**
   * Get the canonical header name for a fuzzy input
   * @param {string} sheetKey - The schema key
   * @param {string} fuzzyHeader - The approximate header name
   * @returns {string|null} - The canonical header
   */
  getCanonicalHeader: function(sheetKey, fuzzyHeader) {
    const schema = this.SCHEMA[sheetKey];
    if (!schema) return null;
    return this.fuzzyMatchHeader(fuzzyHeader, schema.all);
  },

  /**
   * Get column index for a header (using fuzzy matching if needed)
   * @param {string} sheetKey - The schema key
   * @param {Array<string>} actualHeaders - Sheet headers
   * @param {string} canonicalHeader - The header to find
   * @returns {number} - Column index (-1 if not found)
   */
  getColumnIndex: function(sheetKey, actualHeaders, canonicalHeader) {
    const map = this.buildColumnMap(sheetKey, actualHeaders);
    return map[canonicalHeader] !== undefined ? map[canonicalHeader] : -1;
  },

  /**
   * Get all canonical headers for a sheet
   * @param {string} sheetKey - The schema key
   * @returns {Array<string>} - All canonical headers
   */
  getHeaders: function(sheetKey) {
    const schema = this.SCHEMA[sheetKey];
    return schema ? schema.all : [];
  },

  /**
   * Validate that a sheet has all required headers
   * @param {string} sheetKey - The schema key
   * @param {Array<string>} actualHeaders - Headers from the sheet
   * @returns {Object} - Validation result {valid: boolean, missing: Array}
   */
  validateHeaders: function(sheetKey, actualHeaders) {
    const schema = this.SCHEMA[sheetKey];
    if (!schema) {
      return { valid: false, missing: [], error: 'Unknown schema' };
    }

    const map = this.buildColumnMap(sheetKey, actualHeaders);
    const missing = schema.required.filter(req => !(req in map));

    return {
      valid: missing.length === 0,
      missing: missing,
      matched: Object.keys(map),
      unmatched: actualHeaders.filter((_, i) => 
        !Object.values(map).includes(i)
      )
    };
  },

  /**
   * Clear the column map cache
   */
  clearCache: function() {
    this._columnMapCache = {};
  },

  // --- Private helper methods ---

  /**
   * Normalize a string for comparison
   * @private
   */
  _normalize: function(str) {
    if (!str) return '';
    return str
      .toString()
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '') // Remove all non-alphanumeric
      .trim();
  },

  /**
   * Calculate string similarity (0-1)
   * @private
   */
  _similarity: function(a, b) {
    if (a === b) return 1;
    if (!a || !b) return 0;

    const longer = a.length > b.length ? a : b;
    const shorter = a.length > b.length ? b : a;
    
    if (longer.length === 0) return 1;

    const distance = this._levenshteinDistance(longer, shorter);
    return (longer.length - distance) / longer.length;
  },

  /**
   * Levenshtein distance calculation
   * @private
   */
  _levenshteinDistance: function(a, b) {
    const matrix = [];
    
    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }
    
    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }
    
    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        const cost = b[i - 1] === a[j - 1] ? 0 : 1;
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j - 1] + cost
        );
      }
    }
    
    return matrix[b.length][a.length];
  }
};

/**
 * Wrapper for DataLayer that uses SchemaRegistry
 * This ensures all column lookups are fuzzy-matched
 */
const SmartDataLayer = {
  /**
   * Get records with automatic column mapping
   * @param {string} sheetKey - Schema key (e.g., 'PROSPECTS')
   * @returns {Array<Object>} - Records with canonical field names
   */
  getRecords: function(sheetKey) {
    const schema = SchemaRegistry.SCHEMA[sheetKey];
    if (!schema) {
      throw new Error(`Unknown schema: ${sheetKey}`);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(schema.sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet ${schema.sheetName} not found`);
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const actualHeaders = data[0];
    const columnMap = SchemaRegistry.buildColumnMap(sheetKey, actualHeaders);
    
    const records = [];
    
    for (let rowIdx = 1; rowIdx < data.length; rowIdx++) {
      const row = data[rowIdx];
      let hasData = false;
      
      // Check if row has any data
      for (let i = 0; i < row.length; i++) {
        if (row[i] !== '' && row[i] !== null && row[i] !== undefined) {
          hasData = true;
          break;
        }
      }
      
      if (!hasData) continue;
      
      // Build record with canonical field names
      const record = { _rowNumber: rowIdx + 1 };
      
      for (const [canonicalHeader, colIndex] of Object.entries(columnMap)) {
        record[canonicalHeader] = row[colIndex];
      }
      
      records.push(record);
    }
    
    return records;
  },

  /**
   * Add a record with automatic column mapping
   * @param {string} sheetKey - Schema key
   * @param {Object} record - Record with canonical field names
   */
  addRecord: function(sheetKey, record) {
    const schema = SchemaRegistry.SCHEMA[sheetKey];
    if (!schema) {
      throw new Error(`Unknown schema: ${sheetKey}`);
    }

    const lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(30000)) {
        throw new Error('Could not acquire lock');
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(schema.sheetName);
      
      if (!sheet) {
        throw new Error(`Sheet ${schema.sheetName} not found`);
      }

      // Get current headers
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const columnMap = SchemaRegistry.buildColumnMap(sheetKey, headers);
      
      // Build row array
      const maxCol = Math.max(...Object.values(columnMap)) + 1;
      const row = new Array(maxCol).fill('');
      
      for (const [field, value] of Object.entries(record)) {
        const colIndex = columnMap[field];
        if (colIndex !== undefined) {
          row[colIndex] = value;
        }
      }
      
      sheet.appendRow(row);
      return { success: true, message: `Record added to ${schema.sheetName}` };
      
    } catch (error) {
      Logger.log(`Error in addRecord: ${error.message}`);
      throw error;
    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Update a record with automatic column mapping
   * @param {string} sheetKey - Schema key
   * @param {number} rowNumber - Row number to update
   * @param {Object} updates - Fields to update (canonical names)
   */
  updateRecord: function(sheetKey, rowNumber, updates) {
    const schema = SchemaRegistry.SCHEMA[sheetKey];
    if (!schema) {
      throw new Error(`Unknown schema: ${sheetKey}`);
    }

    const lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(30000)) {
        throw new Error('Could not acquire lock');
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(schema.sheetName);
      
      if (!sheet) {
        throw new Error(`Sheet ${schema.sheetName} not found`);
      }

      // Get current headers
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const columnMap = SchemaRegistry.buildColumnMap(sheetKey, headers);
      
      // Update each field
      for (const [field, value] of Object.entries(updates)) {
        const colIndex = columnMap[field];
        if (colIndex !== undefined) {
          sheet.getRange(rowNumber, colIndex + 1).setValue(value);
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
   * Get multiple sheets at once
   * @param {Array<string>} sheetKeys - Array of schema keys
   * @returns {Object} - Object with sheet data
   */
  getMultipleRecords: function(sheetKeys) {
    const result = {};
    sheetKeys.forEach(key => {
      try {
        result[key] = this.getRecords(key);
      } catch (e) {
        Logger.log(`Warning: Could not load ${key}: ${e.message}`);
        result[key] = [];
      }
    });
    return result;
  }
};

/**
 * Validation function to check all sheets against schemas
 * Run this manually to verify data integrity
 */
function validateAllSchemas() {
  const results = [];
  
  for (const [key, schema] of Object.entries(SchemaRegistry.SCHEMA)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(schema.sheetName);
      
      if (!sheet) {
        results.push({ sheet: key, status: 'MISSING', message: 'Sheet not found' });
        continue;
      }
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const validation = SchemaRegistry.validateHeaders(key, headers);
      
      results.push({
        sheet: key,
        status: validation.valid ? 'VALID' : 'INVALID',
        matched: validation.matched.length,
        missing: validation.missing,
        unmatched: validation.unmatched
      });
      
    } catch (e) {
      results.push({ sheet: key, status: 'ERROR', message: e.message });
    }
  }
  
  Logger.log('Schema Validation Results:\n' + JSON.stringify(results, null, 2));
  return results;
}
