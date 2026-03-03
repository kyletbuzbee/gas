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
