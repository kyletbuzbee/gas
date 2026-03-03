/**
 * CRMService.gs
 * Specific business logic for the CRM.
 */

/**
 * Returns a 360-degree view of a specific Company ID.
 * OPTIMIZED: Uses batch fetching for single spreadsheet connection.
 * RELATIONAL FIX: Joins on Company ID only; pulls current name dynamically.
 * @param {string} companyId - The Company ID to look up
 * @returns {Object} - Company profile with associated data
 */
function getCompanyProfile(companyId) {
  try {
    // OPTIMIZED: Fetch all related data in a single spreadsheet connection
    const allData = DataLayer.getMultipleRecords([
      'ACTIVE_CONTAINERS',
      'PROSPECTS',
      'CONTACTS',
      'OUTREACH',
      'SALES'
    ]);

    const activeContainers = allData.ACTIVE_CONTAINERS || [];
    const prospects = allData.PROSPECTS || [];
    const allContacts = allData.CONTACTS || [];
    const allOutreach = allData.OUTREACH || [];
    const allSales = allData.SALES || [];

    // 1. Get Base Profile (Check active first, then prospects)
    let companyProfile = null;
    let isProspect = false;

    const matchActive = activeContainers.find(c => c['Company ID'] === companyId);

    if (matchActive) {
      companyProfile = matchActive;
      companyProfile._type = 'Active';
    } else {
      const matchProspect = prospects.find(p => p['Company ID'] === companyId);
      if (matchProspect) {
        companyProfile = matchProspect;
        companyProfile._type = 'Prospect';
        isProspect = true;
      }
    }

    if (!companyProfile) throw new Error("Company ID not found in Active or Prospects.");

    // RELATIONAL FIX: Use Company ID as the single source of truth
    // Pull the most up-to-date company name from the profile
    const currentCompanyName = companyProfile['Company Name'];

    // 2. Attach associated Contacts (join on Company Name - Contacts sheet doesn't have Company ID)
    // Note: If Contacts sheet has Company ID column, change this to filter by ID instead
    companyProfile.associatedContacts = allContacts.filter(c =>
      c['Company'] === currentCompanyName || c['Company'] === companyId
    );

    // 3. Attach Outreach History - join on Company ID only
    // Company name is pulled dynamically from the profile, not from historical records
    companyProfile.outreachHistory = allOutreach
      .filter(o => o['Company ID'] === companyId)
      .map(o => ({
        ...o,
        // Dynamically override with current company name to handle renames
        'Company': currentCompanyName
      }))
      .sort((a, b) => new Date(b['Visit Date']) - new Date(a['Visit Date']));

    // 4. Attach Sales/Transactions - join on Company ID only
    companyProfile.salesHistory = allSales
      .filter(s => s['Company ID'] === companyId)
      .map(s => ({
        ...s,
        // Dynamically override with current company name to handle renames
        'Company Name': currentCompanyName
      }));

    return companyProfile;

  } catch (error) {
    Logger.log("Error fetching profile: " + error.message);
    return { error: error.message };
  }
}

/**
 * Helper to quickly log a new Outreach/Visit entry
 * RELATIONAL FIX: Looks up company name dynamically by ID to ensure consistency
 * @param {string} companyId - The Company ID
 * @param {string} visitDate - The visit date
 * @param {string} outcome - The outcome of the visit
 * @param {string} notes - Any notes
 * @param {string} contactType - Type of contact (Visit, Phone, Email)
 * @returns {Object} - Result of the add operation
 */
function logOutreachEntry(companyId, visitDate, outcome, notes, contactType) {
  // RELATIONAL FIX: Look up the current company name by ID
  // This ensures the name is always current, even if renamed
  let companyName = lookupCompanyName(companyId, CONFIG.SHEETS.ACTIVE_CONTAINERS);
  if (!companyName) {
    companyName = lookupCompanyName(companyId, CONFIG.SHEETS.PROSPECTS);
  }

  if (!companyName) {
    throw new Error(`Company ID ${companyId} not found in Active or Prospects.`);
  }

  const newEntry = {
    'Outreach ID': 'LID-' + Utilities.getUuid().substring(0, 6).toUpperCase(),
    'Company ID': companyId,
    'Company': companyName, // Always the current name from the source sheet
    'Visit Date': visitDate,
    'Outcome': outcome,
    'Notes': notes,
    'Owner': CONFIG.DEFAULT_OWNER,
    'Contact Type': contactType || 'Visit'
  };

  return DataLayer.addRecord('OUTREACH', newEntry);
}
