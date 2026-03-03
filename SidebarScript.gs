/**
 * Sidebar.gs
 * Functions to show and manage the CRM sidebar in Google Sheets
 */

/**
 * Show the CRM sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('K&L Recycling CRM')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Create custom menu - runs automatically when sheet opens
 */
function onOpen(e) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 K&L CRM')
      .addItem('🚀 Open Dashboard', 'showSidebar')
      .addItem('🔄 Refresh Data', 'refreshData')
      .addSeparator()
      .addItem('➕ Log Outreach', 'showOutreachDialog')
      .addItem('🔍 Company Lookup', 'showCompanySearch')
      .addSeparator()
      .addSubMenu(ui.createMenu('🧹 Data Cleanup')
        .addItem('🔎 Check for Duplicates', 'runDuplicateCheck')
        .addItem('🔀 Merge Duplicates (Interactive)', 'runMergeWizard')
        .addItem('⚡ Batch Merge (Auto)', 'batchMergeObviousDuplicates')
        .addItem('🔢 Generate Next Company ID', 'showNextCompanyId'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🔧 Developer Tools')
        .addItem('✅ Validate Schemas', 'runSchemaValidation')
        .addItem('📝 View Column Mappings', 'showColumnMappings'))
      .addToUi();
    Logger.log('Menu created successfully');
  } catch (error) {
    Logger.log('Error creating menu: ' + error.toString());
  }
}

/**
 * Manual menu creation (run this if menu doesn't appear)
 */
function createMenuManual() {
  onOpen();
  SpreadsheetApp.getUi().alert('Menu created! Check the menu bar.');
}

/**
 * Refresh all data
 */
function refreshData() {
  // Trigger a recalculation
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  
  // Show confirmation
  SpreadsheetApp.getUi().alert('Data refreshed successfully!');
}

/**
 * Show company search dialog
 */
function showCompanySearch() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Inter, sans-serif; padding: 20px; }
        input { width: 100%; padding: 10px; margin: 10px 0; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #3b82f6; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; }
        button:hover { background: #2563eb; }
        .result { padding: 10px; margin: 5px 0; background: #f3f4f6; border-radius: 4px; cursor: pointer; }
        .result:hover { background: #e5e7eb; }
      </style>
    </head>
    <body>
      <h3>🔍 Company Lookup</h3>
      <input type="text" id="search" placeholder="Enter company name or ID...">
      <button onclick="search()">Search</button>
      <div id="results"></div>
      
      <script>
        function search() {
          const term = document.getElementById('search').value;
          google.script.run
            .withSuccessHandler(showResults)
            .searchCompanies(term);
        }
        
        function showResults(companies) {
          const div = document.getElementById('results');
          if (companies.length === 0) {
            div.innerHTML = '<p>No companies found</p>';
            return;
          }
          div.innerHTML = companies.map(c => 
            '<div class="result" onclick="selectCompany(\'' + c.id + '\')">' +
            '<strong>' + c.name + '</strong><br>' +
            '<small>' + c.id + ' • ' + c.city + '</small>' +
            '</div>'
          ).join('');
        }
        
        function selectCompany(id) {
          google.script.run.viewCompanyProfile(id);
          google.script.host.close();
        }
      </script>
    </body>
    </html>
  `)
  .setWidth(400)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Company Lookup');
}

/**
 * OPTIMIZED: Returns lightweight search index for client-side filtering
 * This minimizes data transfer and allows instant search in the sidebar
 * @returns {Array<Object>} - Array of {id, name, city, type} objects
 */
function getCompanySearchIndex() {
  try {
    const allData = DataLayer.getMultipleRecords(['PROSPECTS', 'ACTIVE_CONTAINERS']);

    const prospects = (allData.PROSPECTS || []).map(c => ({
      id: c['Company ID'],
      name: c['Company Name'],
      city: c.City || 'N/A',
      type: 'Prospect'
    }));

    const active = (allData.ACTIVE_CONTAINERS || []).map(c => ({
      id: c['Company ID'],
      name: c['Company Name'],
      city: c.City || 'N/A',
      type: 'Active'
    }));

    return [...prospects, ...active];
  } catch (error) {
    Logger.log('Error getting search index: ' + error.message);
    return [];
  }
}

/**
 * DEPRECATED: Server-side search (kept for backward compatibility)
 * Use getCompanySearchIndex() for new implementations - it returns all data
 * for client-side filtering which is much faster for repeated searches.
 * @param {string} term - Search term
 * @returns {Array<Object>} - Matching companies
 */
function searchCompanies(term) {
  if (!term) return [];

  const all = getCompanySearchIndex();
  const search = term.toLowerCase();

  return all
    .filter(c =>
      c.name?.toLowerCase().includes(search) ||
      c.id?.toLowerCase().includes(search)
    )
    .slice(0, 10);
}

/**
 * View company profile - highlights row in sheet
 */
function viewCompanyProfile(companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Try to find in Prospects sheet
  const prospectsSheet = ss.getSheetByName('Prospects');
  if (prospectsSheet) {
    const data = prospectsSheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === companyId) {
        prospectsSheet.setActiveRange(prospectsSheet.getRange(i + 1, 1));
        return;
      }
    }
  }
  
  // Try Active sheet
  const activeSheet = ss.getSheetByName('Active');
  if (activeSheet) {
    const data = activeSheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === companyId) {
        activeSheet.setActiveRange(activeSheet.getRange(i + 1, 1));
        return;
      }
    }
  }
  
  SpreadsheetApp.getUi().alert('Company not found in sheet');
}

/**
 * Show outreach logging dialog
 */
function showOutreachDialog() {
  const html = HtmlService.createHtmlOutputFromFile('OutreachDialog')
    .setWidth(450)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log Outreach');
}

/**
 * Show next available Company ID
 */
function showNextCompanyId() {
  try {
    const nextId = generateNextCompanyId();
    SpreadsheetApp.getUi().alert(
      'Next Company ID',
      `The next available Company ID is:\n\n${nextId}\n\n` +
      `Copy this ID and use it when creating a new company.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * Run schema validation and show results
 */
function runSchemaValidation() {
  try {
    const results = validateAllSchemas();
    
    let html = '<h3>Schema Validation Results</h3><div style="margin-top:1rem;">';
    
    results.forEach(r => {
      const statusColor = r.status === 'VALID' ? '#16a34a' : r.status === 'MISSING' ? '#6b7280' : '#dc2626';
      html += `<div style="padding:0.75rem;border-bottom:1px solid #eee;">`;
      html += `<strong style="color:${statusColor}">${r.sheet}</strong> - ${r.status}`;
      
      if (r.matched !== undefined) {
        html += `<br><small>Matched: ${r.matched} columns</small>`;
      }
      
      if (r.missing && r.missing.length > 0) {
        html += `<br><small style="color:#dc2626">Missing: ${r.missing.join(', ')}</small>`;
      }
      
      if (r.unmatched && r.unmatched.length > 0) {
        html += `<br><small style="color:#f59e0b">Unmatched: ${r.unmatched.slice(0, 3).join(', ')}${r.unmatched.length > 3 ? '...' : ''}</small>`;
      }
      
      if (r.message) {
        html += `<br><small>${r.message}</small>`;
      }
      
      html += `</div>`;
    });
    
    html += '</div>';
    
    SpreadsheetApp.getUi().alert('Schema Validation', 
      results.filter(r => r.status === 'VALID').length + ' of ' + results.length + ' sheets validated successfully.\n\n' +
      'Check View > Logs for detailed results.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

/**
 * Show column mappings for a selected sheet
 */
function showColumnMappings() {
  const ui = SpreadsheetApp.getUi();
  const sheetNames = Object.keys(SchemaRegistry.SCHEMA);
  
  const response = ui.prompt(
    'View Column Mappings',
    'Enter sheet name (' + sheetNames.join(', ') + '):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const sheetKey = response.getResponseText().toUpperCase().replace(/\s+/g, '_');
  const schema = SchemaRegistry.SCHEMA[sheetKey];
  
  if (!schema) {
    ui.alert('Unknown sheet: ' + sheetKey);
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(schema.sheetName);
    
    if (!sheet) {
      ui.alert('Sheet not found: ' + schema.sheetName);
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnMap = SchemaRegistry.buildColumnMap(sheetKey, headers);
    
    let html = `<h3>Column Mappings: ${schema.sheetName}</h3>`;
    html += '<table style="width:100%;border-collapse:collapse;margin-top:1rem;">';
    html += '<tr style="background:#f3f4f6;"><th style="padding:0.5rem;text-align:left;border:1px solid #ddd;">Canonical Name</th>';
    html += '<th style="padding:0.5rem;text-align:left;border:1px solid #ddd;">Actual Header</th>';
    html += '<th style="padding:0.5rem;text-align:left;border:1px solid #ddd;">Column</th></tr>';
    
    for (const [canonical, index] of Object.entries(columnMap).sort((a, b) => a[1] - b[1])) {
      const actualHeader = headers[index];
      const matchClass = canonical === actualHeader ? 'color:#16a34a' : 'color:#2563eb';
      html += `<tr><td style="padding:0.5rem;border:1px solid #eee;">${canonical}</td>`;
      html += `<td style="padding:0.5rem;border:1px solid #eee;${matchClass}">${actualHeader}</td>`;
      html += `<td style="padding:0.5rem;border:1px solid #eee;">${index + 1}</td></tr>`;
    }
    
    // Show unmatched headers
    const unmatched = headers.filter((_, i) => !Object.values(columnMap).includes(i));
    if (unmatched.length > 0) {
      html += '</table><h4 style="margin-top:1.5rem;color:#6b7280;">Unmatched Headers:</h4><ul>';
      unmatched.forEach(h => {
        html += `<li style="color:#6b7280;">${h}</li>`;
      });
      html += '</ul>';
    }
    
    ui.alert('Column Mappings', 'View > Logs for detailed mapping table', ui.ButtonSet.OK);
    Logger.log(html);
    
  } catch (error) {
    ui.alert('Error: ' + error.message);
  }
}

/**
 * Get dashboard stats for quick view
 * OPTIMIZED: Uses getMultipleRecords for single spreadsheet connection
 * Excludes won accounts from prospect counts
 */
function getDashboardStats() {
  try {
    const allData = DataLayer.getMultipleRecords([
      'PROSPECTS', 'ACTIVE_CONTAINERS', 'OUTREACH', 'SALES', 'ACCOUNTS'
    ]);

    const prospects = allData.PROSPECTS || [];
    const active = allData.ACTIVE_CONTAINERS || [];
    const outreach = allData.OUTREACH || [];
    const sales = allData.SALES || [];
    const accounts = allData.ACCOUNTS || [];

    // Get won company IDs from outreach and accounts
    const wonIds = new Set();
    outreach.forEach(o => {
      if (o.Stage === 'Won' || (o['Outcome'] && o['Outcome'].toLowerCase().includes('won'))) {
        wonIds.add(o['Company ID']);
      }
    });

    // Also mark accounts as won
    accounts.forEach(a => {
      const companyName = a['Company Name'];
      [...prospects, ...active].forEach(p => {
        if (p['Company Name'] === companyName) {
          wonIds.add(p['Company ID']);
        }
      });
    });

    // Filter prospects to exclude won
    const activeProspects = prospects.filter(p => !wonIds.has(p['Company ID']));

    const thisMonth = new Date().getMonth();
    const thisYear = new Date().getFullYear();

    // Revenue from SALES sheet only
    const totalRevenue = sales.reduce((sum, s) => sum + (parseFloat(s['Payment Amount']) || 0), 0);

    return {
      totalProspects: activeProspects.length,
      totalActive: active.length,
      totalAccounts: accounts.length,
      outreachThisMonth: outreach.filter(o => {
        const d = new Date(o['Visit Date']);
        return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
      }).length,
      totalRevenue: totalRevenue,
      hotLeads: activeProspects.filter(p => p['UrgencyBand'] === 'High').length,
      conversionRate: outreach.length > 0
        ? ((outreach.filter(o => o.Stage === 'Won').length / outreach.length) * 100).toFixed(1)
        : 0
    };
  } catch (error) {
    Logger.log('Error getting stats: ' + error.message);
    return {};
  }
}
