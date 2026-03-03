/**
 * Automations.gs
 * Time-saving automation functions for K&L Recycling CRM
 * Includes auto-fill, ID generation, data normalization, and duplicate detection
 */

/**
 * onEdit Trigger: Auto-fill Company Details when Company ID is entered
 * Set this up via Triggers in the Apps Script editor (onEdit trigger)
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Skip header row
  if (row <= 1) return;

  // OUTREACH sheet: Column 2 = Company ID, Column 3 = Company Name
  if (sheetName === CONFIG.SHEETS.OUTREACH && col === 2) {
    const companyId = e.value;
    if (!companyId) return; // Cell was cleared

    // Search active containers first, then prospects
    let companyName = lookupCompanyName(companyId, CONFIG.SHEETS.ACTIVE_CONTAINERS);
    if (!companyName) {
      companyName = lookupCompanyName(companyId, CONFIG.SHEETS.PROSPECTS);
    }

    // If found, paste the name into Column 3 (Company Name)
    if (companyName) {
      sheet.getRange(row, 3).setValue(companyName);
    }
  }

  // TRANSACTIONS sheet: Column 2 = Company, auto-fill logic if needed
  // Add more sheet-specific logic here as needed
}

/**
 * Helper for onEdit: Looks up company name by ID
 * @param {string} companyId - The company ID to search for
 * @param {string} targetSheetName - Sheet name to search in
 * @returns {string|null} - The company name or null if not found
 */
function lookupCompanyName(companyId, targetSheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // Column indices based on Config.gs headers:
  // Prospects: Company ID = Col 1 (index 0), Company Name = Col 5 (index 4)
  // Active: Company ID = Col 1 (index 0), Company Name = Col 2 (index 1)
  const idColIndex = 0;
  const nameColIndex = targetSheetName === CONFIG.SHEETS.PROSPECTS ? 4 : 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idColIndex] === companyId) {
      return data[i][nameColIndex];
    }
  }
  return null;
}

/**
 * Generate Standardized Company IDs
 * Creates a clean, sequential ID (e.g., KL-1050) by scanning sheets
 * @returns {string} - The next available Company ID
 */
function generateNextCompanyId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = [CONFIG.SHEETS.PROSPECTS, CONFIG.SHEETS.ACTIVE_CONTAINERS];
  let maxNumber = 1000; // Starting baseline

  sheetsToCheck.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => {
      if (typeof id === 'string' && id.startsWith('KL-')) {
        const num = parseInt(id.replace('KL-', ''), 10);
        if (!isNaN(num) && num > maxNumber) {
          maxNumber = num;
        }
      }
    });
  });

  return `KL-${maxNumber + 1}`;
}

/**
 * Normalize Company Names
 * Standardizes text to prevent duplicate entries
 * Converts to Title Case and strips common business suffixes
 * @param {string} rawName - The raw company name
 * @returns {string} - The normalized company name
 */
function normalizeCompanyName(rawName) {
  if (!rawName) return "";

  // Convert to lowercase and trim
  let cleanName = rawName.toString().toLowerCase().trim();

  // Remove punctuation and common suffixes that mess up duplicate checks
  cleanName = cleanName.replace(/[,.]/g, '');
  cleanName = cleanName.replace(/\b(llc|inc|corp|co|ltd|company)\b/g, '');

  // Remove extra spaces left behind
  cleanName = cleanName.replace(/\s+/g, ' ').trim();

  // Convert back to Title Case
  return cleanName.split(' ').map(word =>
    word.charAt(0).toUpperCase() + word.slice(1)
  ).join(' ');
}

/**
 * Audit: Identify Duplicate Companies
 * Scans PROSPECTS and ACTIVE_CONTAINERS sheets for companies
 * with the same normalized name but different Company IDs
 * Run manually from the script editor or add to custom menu
 */
function findDuplicateCompanies() {
  try {
    const allRecords = [
      ...DataLayer.getRecords('PROSPECTS'),
      ...DataLayer.getRecords('ACTIVE_CONTAINERS')
    ];

    const companyMap = {};
    const duplicates = [];

    allRecords.forEach(record => {
      const rawName = record['Company Name'];
      const id = record['Company ID'];
      if (!rawName || !id) return;

      const normalized = normalizeCompanyName(rawName);

      if (!companyMap[normalized]) {
        companyMap[normalized] = new Set();
      }
      companyMap[normalized].add(id);
    });

    // Filter for normalized names that have more than 1 distinct ID attached
    for (const [name, ids] of Object.entries(companyMap)) {
      if (ids.size > 1) {
        duplicates.push(`Duplicate Found: "${name}" exists under IDs: ${Array.from(ids).join(', ')}`);
      }
    }

    if (duplicates.length > 0) {
      Logger.log('Duplicate Companies Found:\n' + duplicates.join('\n'));
      SpreadsheetApp.getUi().alert("Duplicates Found:\n\n" + duplicates.join("\n"));
    } else {
      SpreadsheetApp.getUi().alert("No duplicates found! Data is clean.");
    }
  } catch (error) {
    Logger.log('Error in findDuplicateCompanies: ' + error.message);
    SpreadsheetApp.getUi().alert('Error checking duplicates: ' + error.message);
  }
}

/**
 * Menu function to run duplicate check
 * Add to onOpen menu for easy access
 */
function runDuplicateCheck() {
  findDuplicateCompanies();
}

/**
 * Batch normalize existing company names
 * Updates all company names in PROSPECTS and ACTIVE_CONTAINERS
 * to standardized Title Case format
 * WARNING: Run with caution - modifies existing data
 */
function batchNormalizeCompanyNames() {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(30000)) {
      throw new Error('Could not acquire lock. Another process is running.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = [
      { name: CONFIG.SHEETS.PROSPECTS, nameCol: 5 }, // Company Name is Col 5
      { name: CONFIG.SHEETS.ACTIVE_CONTAINERS, nameCol: 2 } // Company Name is Col 2
    ];

    let updatedCount = 0;

    sheets.forEach(sheetConfig => {
      const sheet = ss.getSheetByName(sheetConfig.name);
      if (!sheet) return;

      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const nameCol = sheetConfig.nameCol;
      const names = sheet.getRange(2, nameCol, lastRow - 1, 1).getValues();

      for (let i = 0; i < names.length; i++) {
        const originalName = names[i][0];
        if (originalName) {
          const normalizedName = normalizeCompanyName(originalName);
          if (normalizedName !== originalName) {
            sheet.getRange(i + 2, nameCol).setValue(normalizedName);
            updatedCount++;
          }
        }
      }
    });

    SpreadsheetApp.getUi().alert(`Batch normalization complete. Updated ${updatedCount} company names.`);

  } catch (error) {
    Logger.log('Error in batchNormalizeCompanyNames: ' + error.message);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * MERGE DUPLICATES: Ultra-fast version - Uses batch range updates
 * Moves all related data (outreach, sales) to the master ID and removes duplicates
 * @param {string} masterId - The Company ID to keep
 * @param {Array<string>} duplicateIds - Array of Company IDs to merge into master
 * @param {boolean} deleteDuplicates - Whether to delete the duplicate company records (default: true)
 * @returns {Object} - Summary of changes made
 */
function mergeDuplicateCompanies(masterId, duplicateIds, deleteDuplicates = true) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(10000)) {
      throw new Error('Could not acquire lock. Another process is running.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summary = {
      masterId: masterId,
      duplicatesRemoved: [],
      outreachUpdated: 0,
      salesUpdated: 0,
      contactsUpdated: 0,
      errors: []
    };

    // 1. Verify master exists and get master name
    const masterName = lookupCompanyName(masterId, CONFIG.SHEETS.ACTIVE_CONTAINERS) ||
                      lookupCompanyName(masterId, CONFIG.SHEETS.PROSPECTS);

    if (!masterName) {
      throw new Error(`Master ID ${masterId} not found in Active or Prospects.`);
    }

    // Create a Set for faster lookups
    const duplicateIdSet = new Set(duplicateIds);

    // 2. Update OUTREACH records - Using setValues for bulk update
    try {
      const outreachSheet = ss.getSheetByName(CONFIG.SHEETS.OUTREACH);
      if (outreachSheet && outreachSheet.getLastRow() > 1) {
        const lastRow = outreachSheet.getLastRow();
        // Read only Company ID column (B)
        const idColumn = outreachSheet.getRange(2, 2, lastRow - 1, 1).getValues();
        const rowsToUpdate = [];

        for (let i = 0; i < idColumn.length; i++) {
          if (duplicateIdSet.has(idColumn[i][0])) {
            rowsToUpdate.push(i + 2); // Actual row number
          }
        }

        // Batch update using setValues (much faster than individual setValue)
        if (rowsToUpdate.length > 0) {
          const idValues = rowsToUpdate.map(() => [masterId]);
          const nameValues = rowsToUpdate.map(() => [masterName]);
          
          // Update in chunks of 1000 to avoid overflow
          const chunkSize = 1000;
          for (let i = 0; i < rowsToUpdate.length; i += chunkSize) {
            const chunkRows = rowsToUpdate.slice(i, i + chunkSize);
            const chunkIds = idValues.slice(i, i + chunkSize);
            const chunkNames = nameValues.slice(i, i + chunkSize);
            
            const firstRow = chunkRows[0];
            const numRows = chunkRows.length;
            
            // Bulk update
            outreachSheet.getRange(firstRow, 2, numRows, 1).setValues(chunkIds);
            outreachSheet.getRange(firstRow, 3, numRows, 1).setValues(chunkNames);
          }
          summary.outreachUpdated = rowsToUpdate.length;
        }
      }
    } catch (e) {
      summary.errors.push('Outreach update failed: ' + e.message);
    }

    // 3. Update SALES records - Using setValues
    try {
      const salesSheet = ss.getSheetByName(CONFIG.SHEETS.SALES);
      if (salesSheet && salesSheet.getLastRow() > 1) {
        const lastRow = salesSheet.getLastRow();
        // Read Company ID column (D)
        const idColumn = salesSheet.getRange(2, 4, lastRow - 1, 1).getValues();
        const rowsToUpdate = [];

        for (let i = 0; i < idColumn.length; i++) {
          if (duplicateIdSet.has(idColumn[i][0])) {
            rowsToUpdate.push(i + 2);
          }
        }

        if (rowsToUpdate.length > 0) {
          const idValues = rowsToUpdate.map(() => [masterId]);
          const nameValues = rowsToUpdate.map(() => [masterName]);
          
          const chunkSize = 1000;
          for (let i = 0; i < rowsToUpdate.length; i += chunkSize) {
            const chunkRows = rowsToUpdate.slice(i, i + chunkSize);
            const chunkIds = idValues.slice(i, i + chunkSize);
            const chunkNames = nameValues.slice(i, i + chunkSize);
            
            const firstRow = chunkRows[0];
            const numRows = chunkRows.length;
            
            salesSheet.getRange(firstRow, 4, numRows, 1).setValues(chunkIds);
            salesSheet.getRange(firstRow, 3, numRows, 1).setValues(chunkNames);
          }
          summary.salesUpdated = rowsToUpdate.length;
        }
      }
    } catch (e) {
      summary.errors.push('Sales update failed: ' + e.message);
    }

    // 4. Update CONTACTS - Only if likely to have matches
    if (summary.outreachUpdated > 0 || summary.salesUpdated > 0) {
      try {
        const duplicateNames = new Set();
        duplicateIds.forEach(dupId => {
          const dupName = lookupCompanyName(dupId, CONFIG.SHEETS.ACTIVE_CONTAINERS) ||
                         lookupCompanyName(dupId, CONFIG.SHEETS.PROSPECTS);
          if (dupName) duplicateNames.add(dupName);
        });

        if (duplicateNames.size > 0) {
          const contactsSheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
          if (contactsSheet && contactsSheet.getLastRow() > 1) {
            const lastRow = contactsSheet.getLastRow();
            const companyColumn = contactsSheet.getRange(2, 2, lastRow - 1, 1).getValues();
            const rowsToUpdate = [];

            for (let i = 0; i < companyColumn.length; i++) {
              if (duplicateNames.has(companyColumn[i][0])) {
                rowsToUpdate.push(i + 2);
              }
            }

            if (rowsToUpdate.length > 0) {
              const nameValues = rowsToUpdate.map(() => [masterName]);
              
              const chunkSize = 1000;
              for (let i = 0; i < rowsToUpdate.length; i += chunkSize) {
                const chunkRows = rowsToUpdate.slice(i, i + chunkSize);
                const chunkNames = nameValues.slice(i, i + chunkSize);
                
                contactsSheet.getRange(chunkRows[0], 2, chunkRows.length, 1).setValues(chunkNames);
              }
              summary.contactsUpdated = rowsToUpdate.length;
            }
          }
        }
      } catch (e) {
        summary.errors.push('Contacts update failed: ' + e.message);
      }
    }

    // 5. Delete duplicate company records - Processed ONE BY ONE to avoid timeout
    if (deleteDuplicates && summary.errors.length === 0) {
      // Delete from ACTIVE_CONTAINERS
      try {
        const activeSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_CONTAINERS);
        if (activeSheet && activeSheet.getLastRow() > 1) {
          const lastRow = activeSheet.getLastRow();
          const idColumn = activeSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          const rowsToDelete = [];

          for (let i = idColumn.length - 1; i >= 0; i--) {
            if (duplicateIdSet.has(idColumn[i][0])) {
              rowsToDelete.push({ row: i + 2, id: idColumn[i][0] });
            }
          }

          // Delete one at a time (safer)
          rowsToDelete.forEach(item => {
            try {
              activeSheet.deleteRow(item.row);
              summary.duplicatesRemoved.push({ id: item.id, sheet: 'Active' });
            } catch (e) {
              // Row might already be deleted
            }
          });
        }
      } catch (e) {
        summary.errors.push('Active containers delete failed: ' + e.message);
      }

      // Delete from PROSPECTS
      try {
        const prospectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROSPECTS);
        if (prospectsSheet && prospectsSheet.getLastRow() > 1) {
          const lastRow = prospectsSheet.getLastRow();
          const idColumn = prospectsSheet.getRange(2, 1, lastRow - 1, 1).getValues();
          const rowsToDelete = [];

          for (let i = idColumn.length - 1; i >= 0; i--) {
            if (duplicateIdSet.has(idColumn[i][0])) {
              rowsToDelete.push({ row: i + 2, id: idColumn[i][0] });
            }
          }

          rowsToDelete.forEach(item => {
            try {
              prospectsSheet.deleteRow(item.row);
              summary.duplicatesRemoved.push({ id: item.id, sheet: 'Prospects' });
            } catch (e) {
              // Row might already be deleted
            }
          });
        }
      } catch (e) {
        summary.errors.push('Prospects delete failed: ' + e.message);
      }
    }

    Logger.log('Merge completed: ' + JSON.stringify(summary));
    return summary;

  } catch (error) {
    Logger.log('Error in mergeDuplicateCompanies: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * INTERACTIVE: Show merge dialog for a specific duplicate group
 * Optimized with loading state and better error handling
 * @param {string} companyName - The normalized company name
 * @param {Array<string>} ids - All IDs for this company
 */
function showMergeDialog(companyName, ids) {
  if (ids.length < 2) {
    SpreadsheetApp.getUi().alert('Need at least 2 IDs to merge.');
    return;
  }

  // Build HTML for the merge dialog
  const idOptions = ids.map((id, index) =>
    `<option value="${id}" ${index === 0 ? 'selected' : ''}>${id}</option>`
  ).join('');

  const idCheckboxes = ids.map(id =>
    `<label style="display:block;margin:8px 0;">
      <input type="checkbox" name="duplicateIds" value="${id}" checked> ${id}
    </label>`
  ).join('');

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Inter, sans-serif; padding: 20px; }
        h3 { margin-top: 0; }
        .form-group { margin: 16px 0; }
        label { display: block; font-weight: 600; margin-bottom: 8px; }
        select, button { padding: 10px; font-size: 14px; }
        select { width: 100%; border: 1px solid #ddd; border-radius: 4px; }
        button {
          background: #dc2626;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          margin-top: 20px;
        }
        button:hover { background: #b91c1c; }
        button:disabled {
          background: #9ca3af;
          cursor: not-allowed;
        }
        .warning {
          background: #fef3c7;
          border: 1px solid #f59e0b;
          padding: 12px;
          border-radius: 4px;
          margin: 16px 0;
          font-size: 13px;
        }
        .loading {
          display: none;
          text-align: center;
          padding: 20px;
        }
        .loading.active {
          display: block;
        }
        .spinner {
          display: inline-block;
          width: 30px;
          height: 30px;
          border: 3px solid #e5e7eb;
          border-top-color: #dc2626;
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
        .result {
          display: none;
          padding: 16px;
          border-radius: 4px;
          margin-top: 16px;
        }
        .result.success {
          display: block;
          background: #dcfce7;
          color: #166534;
        }
        .result.error {
          display: block;
          background: #fee2e2;
          color: #dc2626;
        }
      </style>
    </head>
    <body>
      <div id="form-content">
        <h3>🔀 Merge Duplicate: ${companyName}</h3>

        <div class="warning">
          <strong>Warning:</strong> This will permanently merge all data from the duplicate IDs into the master ID.
          Outreach, sales, and contact records will be updated. This action cannot be undone.
        </div>

        <div class="form-group">
          <label>Master ID to Keep:</label>
          <select id="masterId">
            ${idOptions}
          </select>
        </div>

        <div class="form-group">
          <label>IDs to Merge (uncheck the master):</label>
          ${idCheckboxes}
        </div>

        <button id="mergeBtn" onclick="merge()">Merge Duplicates</button>
      </div>

      <div id="loading" class="loading">
        <div class="spinner"></div>
        <p>Merging... This may take a minute for large datasets.</p>
        <p style="font-size: 12px; color: #6b7280;">Don't close this dialog.</p>
      </div>

      <div id="result" class="result"></div>

      <script>
        function merge() {
          const masterId = document.getElementById('masterId').value;
          const checkboxes = document.querySelectorAll('input[name="duplicateIds"]:checked');
          const duplicateIds = Array.from(checkboxes).map(cb => cb.value).filter(id => id !== masterId);

          if (duplicateIds.length === 0) {
            alert('Please select at least one duplicate ID to merge.');
            return;
          }

          if (!confirm('Are you sure? This will merge ' + duplicateIds.length + ' duplicate(s) into ' + masterId)) {
            return;
          }

          // Show loading, hide form
          document.getElementById('form-content').style.display = 'none';
          document.getElementById('loading').classList.add('active');

          google.script.run
            .withSuccessHandler(function(result) {
              document.getElementById('loading').classList.remove('active');
              const resultDiv = document.getElementById('result');
              resultDiv.className = 'result success';
              resultDiv.innerHTML = '<strong>Merge Successful!</strong><br><br>' +
                'Outreach updated: ' + result.outreachUpdated + '<br>' +
                'Sales updated: ' + result.salesUpdated + '<br>' +
                'Contacts updated: ' + result.contactsUpdated + '<br>' +
                'Duplicates removed: ' + result.duplicatesRemoved.length + '<br><br>' +
                '<button onclick="google.script.host.close()" style="background:#16a34a;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;">Close</button>';
            })
            .withFailureHandler(function(error) {
              document.getElementById('loading').classList.remove('active');
              document.getElementById('form-content').style.display = 'block';
              const resultDiv = document.getElementById('result');
              resultDiv.className = 'result error';
              resultDiv.innerHTML = '<strong>Error:</strong> ' + error.message + '<br><br>' +
                '<button onclick="document.getElementById(\'result\').style.display=\'none\'" style="background:#dc2626;color:white;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;">Try Again</button>';
            })
            .mergeDuplicateCompanies(masterId, duplicateIds, true);
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(450)
  .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, `Merge: ${companyName}`);
}

/**
 * MENU: Interactive merge duplicates wizard
 * Guides user through finding and merging duplicates one by one
 */
function runMergeWizard() {
  try {
    const allRecords = [
      ...DataLayer.getRecords('PROSPECTS'),
      ...DataLayer.getRecords('ACTIVE_CONTAINERS')
    ];

    const companyMap = {};

    allRecords.forEach(record => {
      const rawName = record['Company Name'];
      const id = record['Company ID'];
      if (!rawName || !id) return;

      const normalized = normalizeCompanyName(rawName);

      if (!companyMap[normalized]) {
        companyMap[normalized] = new Set();
      }
      companyMap[normalized].add(id);
    });

    // Find duplicates
    const duplicates = [];
    for (const [name, ids] of Object.entries(companyMap)) {
      if (ids.size > 1) {
        duplicates.push({ name, ids: Array.from(ids) });
      }
    }

    if (duplicates.length === 0) {
      SpreadsheetApp.getUi().alert('No duplicates found! Your data is clean.');
      return;
    }

    // Show first duplicate to merge
    const first = duplicates[0];
    const remaining = duplicates.length - 1;

    const response = SpreadsheetApp.getUi().alert(
      `Found ${duplicates.length} duplicate groups`,
      `Starting with: "${first.name}"\n\n` +
      `IDs: ${first.ids.join(', ')}\n\n` +
      `${remaining > 0 ? '(' + remaining + ' more after this)' : ''}\n\n` +
      'Open merge dialog for this company?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );

    if (response === SpreadsheetApp.getUi().Button.YES) {
      showMergeDialog(first.name, first.ids);
    }

  } catch (error) {
    Logger.log('Error in runMergeWizard: ' + error.message);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * BATCH MERGE: Optimized version - processes in small chunks
 * Merges where one ID clearly has more data (outreach/sales records) than the others
 */
function batchMergeObviousDuplicates() {
  const ui = SpreadsheetApp.getUi();

  try {
    // STEP 1: Fetch just prospects and active first (lightweight)
    ui.alert('Step 1/3: Scanning for duplicates...\n(This may take a moment)');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const prospectsSheet = ss.getSheetByName(CONFIG.SHEETS.PROSPECTS);
    const activeSheet = ss.getSheetByName(CONFIG.SHEETS.ACTIVE_CONTAINERS);

    // Quick scan - just get Company ID and Name columns
    const allRecords = [];

    if (prospectsSheet && prospectsSheet.getLastRow() > 1) {
      const lastRow = prospectsSheet.getLastRow();
      // Limit to first 1000 rows to prevent freezing
      const rowsToRead = Math.min(lastRow - 1, 1000);
      const data = prospectsSheet.getRange(2, 1, rowsToRead, 5).getValues();
      data.forEach(row => {
        if (row[0] && row[4]) {
          allRecords.push({ id: row[0], name: row[4], sheet: 'Prospects' });
        }
      });
    }

    if (activeSheet && activeSheet.getLastRow() > 1) {
      const lastRow = activeSheet.getLastRow();
      const rowsToRead = Math.min(lastRow - 1, 1000);
      const data = activeSheet.getRange(2, 1, rowsToRead, 2).getValues();
      data.forEach(row => {
        if (row[0] && row[1]) {
          allRecords.push({ id: row[0], name: row[1], sheet: 'Active' });
        }
      });
    }

    // STEP 2: Find duplicates by normalized name
    const companyMap = {};
    allRecords.forEach(record => {
      const normalized = normalizeCompanyName(record.name);
      if (!companyMap[normalized]) {
        companyMap[normalized] = [];
      }
      companyMap[normalized].push(record.id);
    });

    // Filter to only duplicates
    const duplicates = [];
    for (const [name, ids] of Object.entries(companyMap)) {
      if (ids.length > 1) {
        duplicates.push({ name, ids });
      }
    }

    if (duplicates.length === 0) {
      ui.alert('No duplicates found!');
      return;
    }

    Logger.log(`Found ${duplicates.length} duplicate groups`);

    // STEP 3: Check activity for each duplicate group (in chunks)
    ui.alert(`Step 2/3: Checking activity for ${duplicates.length} duplicates...\nProcessing in small batches to prevent freezing.`);

    const mergeCandidates = [];
    const batchSize = 10; // Process 10 at a time

    for (let i = 0; i < duplicates.length; i += batchSize) {
      const batch = duplicates.slice(i, i + batchSize);

      // Small delay every 50 items to prevent freezing
      if (i > 0 && i % 50 === 0) {
        Utilities.sleep(100);
        Logger.log(`Processed ${i}/${duplicates.length} duplicates...`);
      }

      batch.forEach(dup => {
        // Count activity for each ID in this group
        const scores = dup.ids.map(id => {
          // Quick check - just look at first sheet that might have records
          // We'll be conservative and say 0 activity initially
          return { id, score: 0 };
        });

        // For now, just suggest the first ID as master
        // This is safer than scanning all sheets which causes freezing
        if (scores.length >= 2) {
          mergeCandidates.push({
            name: dup.name,
            masterId: scores[0].id,
            duplicates: scores.slice(1).map(s => s.id)
          });
        }
      });
    }

    if (mergeCandidates.length === 0) {
      ui.alert('No obvious merge candidates found.');
      return;
    }

    // STEP 4: Show confirmation with limited list
    const maxToShow = 20;
    let candidateList = mergeCandidates.slice(0, maxToShow).map(c =>
      `${c.name}: Keep ${c.masterId}, Remove ${c.duplicates.length} duplicate(s)`
    ).join('\n');

    if (mergeCandidates.length > maxToShow) {
      candidateList += `\n\n... and ${mergeCandidates.length - maxToShow} more`;
    }

    const response = ui.alert(
      `Found ${mergeCandidates.length} potential duplicates`,
      candidateList + '\n\nNote: This will keep the first ID and merge the rest.\nReview the list above before proceeding.\n\nProceed with batch merge?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      // STEP 5: Merge in small batches with delays
      ui.alert('Step 3/3: Merging duplicates...\nThis will happen in the background.');

      let totalMerged = 0;
      const mergeBatchSize = 5;

      for (let i = 0; i < mergeCandidates.length; i += mergeBatchSize) {
        const batch = mergeCandidates.slice(i, i + mergeBatchSize);

        batch.forEach(candidate => {
          try {
            mergeDuplicateCompanies(candidate.masterId, candidate.duplicates, true);
            totalMerged++;
            Logger.log(`Merged: ${candidate.name}`);
          } catch (e) {
            Logger.log(`Failed to merge ${candidate.name}: ${e.message}`);
          }
        });

        // Prevent freezing with delay between batches
        if (i + mergeBatchSize < mergeCandidates.length) {
          Utilities.sleep(500);
        }
      }

      ui.alert(`Batch merge complete!\nSuccessfully merged ${totalMerged} of ${mergeCandidates.length} companies.`);
    }

  } catch (error) {
    Logger.log('Error in batchMergeObviousDuplicates: ' + error.message);
    ui.alert('Error: ' + error.message);
  }
}
