/**
 * WebApp.gs
 * Entry point for the CRM web application
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('K&L Recycling CRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns all data needed for the dashboard
 * Called from client-side via google.script.run
 * OPTIMIZED: Fetches in prioritized chunks to prevent timeout
 */
function getAllData() {
  try {
    Logger.log('Starting getAllData...');
    
    // PRIORITY 1: Essential data for dashboard (fetch first)
    const essentialData = DataLayer.getMultipleRecords([
      'PROSPECTS', 'ACTIVE_CONTAINERS', 'OUTREACH', 'SALES'
    ]);
    
    const prospects = essentialData.PROSPECTS || [];
    const active = essentialData.ACTIVE_CONTAINERS || [];
    const outreach = essentialData.OUTREACH || [];
    const sales = essentialData.SALES || [];
    
    Logger.log(`Loaded essential: ${prospects.length} prospects, ${active.length} active, ${outreach.length} outreach, ${sales.length} sales`);
    
    // PRIORITY 2: Secondary data (accounts, contacts)
    let accounts = [], contacts = [], transactions = [];
    try {
      const secondaryData = DataLayer.getMultipleRecords([
        'ACCOUNTS', 'CONTACTS', 'TRANSACTIONS'
      ]);
      accounts = secondaryData.ACCOUNTS || [];
      contacts = secondaryData.CONTACTS || [];
      transactions = secondaryData.TRANSACTIONS || [];
      Logger.log(`Loaded secondary: ${accounts.length} accounts, ${contacts.length} contacts, ${transactions.length} transactions`);
    } catch (e) {
      Logger.log('Warning: Could not load secondary data: ' + e.message);
    }
    
    const rawData = {
      prospects: prospects,
      active: active,
      outreach: outreach,
      sales: sales,
      contacts: contacts,
      transactions: transactions,
      accounts: accounts
    };

    // Convert Dates to strings
    return sanitizeForFrontend(rawData);
    
  } catch (error) {
    Logger.log('Critical error in getAllData: ' + error.toString());
    return { error: error.toString(), message: 'Failed to load data. Please try again.' };
  }
}

/**
 * Helper function to safely convert Dates to strings before sending to HTML
 */
function sanitizeForFrontend(dataObj) {
  const timezone = Session.getScriptTimeZone();
  
  // Helper to process arrays of row objects
  const processArray = (arr) => {
    return arr.map(row => {
      const newRow = {};
      for (let key in row) {
        if (row[key] instanceof Date) {
          // Convert Date to a clean, readable string format
          newRow[key] = Utilities.formatDate(row[key], timezone, "MM/dd/yyyy");
        } else {
          newRow[key] = row[key];
        }
      }
      return newRow;
    });
  };

  return {
    prospects: processArray(dataObj.prospects || []),
    active: processArray(dataObj.active || []),
    outreach: processArray(dataObj.outreach || []),
    sales: processArray(dataObj.sales || []),
    contacts: processArray(dataObj.contacts || []),
    transactions: processArray(dataObj.transactions || []),
    accounts: processArray(dataObj.accounts || [])
  };
}

/**
 * Get company profile with related data
 */
function getCompanyProfileData(companyId) {
  return getCompanyProfile(companyId);
}

/**
 * Add new outreach entry
 * Updated to use Company ID as single source of truth
 */
function addOutreachEntry(entry) {
  return logOutreachEntry(
    entry.companyId,
    entry.visitDate,
    entry.outcome,
    entry.notes,
    entry.contactType
  );
}

/**
 * Get material prices
 */
function getPrices() {
  try {
    return DataLayer.getRecords('PRICES');
  } catch (error) {
    Logger.log('Error loading prices: ' + error.message);
    return [];
  }
}

/**
 * Get dashboard summary stats
 * OPTIMIZED: Uses getMultipleRecords for single spreadsheet connection
 */
function getDashboardStats() {
  try {
    const allData = DataLayer.getMultipleRecords(['PROSPECTS', 'ACTIVE_CONTAINERS', 'OUTREACH', 'SALES']);

    const prospects = allData.PROSPECTS || [];
    const active = allData.ACTIVE_CONTAINERS || [];
    const outreach = allData.OUTREACH || [];
    const sales = allData.SALES || [];

    const thisMonth = new Date().getMonth();
    const thisYear = new Date().getFullYear();

    return {
      totalProspects: prospects.length,
      totalActive: active.length,
      outreachThisMonth: outreach.filter(o => {
        const d = new Date(o['Visit Date']);
        return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
      }).length,
      totalSales: sales.reduce((sum, s) => sum + (parseFloat(s['Payment Amount']) || 0), 0)
    };
  } catch (error) {
    Logger.log('Error getting stats: ' + error.message);
    return {};
  }
}
