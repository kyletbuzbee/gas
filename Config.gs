/**
 * K&L Recycling CRM - Configuration (Config.gs)
 * Centralized constants for Sheet names and Column Headers.
 * UPDATED: Aligned with all actual CSV file headers in gas/ folder
 */

const CONFIG = {
  // Sheet Names - All sheets found in gas/ folder
  SHEETS: {
    OUTREACH: 'Outreach',
    PROSPECTS: 'Prospects',
    RAW_PROSPECTS: 'Raw Prospects',
    STALE_PROSPECTS: 'Stale Prospects',
    ACCOUNTS: 'Accounts',
    CONTACTS: 'Contacts',
    DASHBOARD: 'Dashboard',
    METRICS: 'MetricsHistory',
    SYSTEM_LOG: 'System_OpsLog',
    SETTINGS: 'Settings',
    PRICES: 'Prices',
    SALES: 'Sales',
    ACTIVE_CONTAINERS: 'Active',
    TRANSACTIONS: 'Transactions',
    SYSTEM_SCHEMA: 'System_Schema'
  },

  APP_TITLE: 'K&L Recycling CRM',
  
  // Header Definitions (Must match Sheet Headers EXACTLY)
  HEADERS: {
    OUTREACH: [
      'Outreach ID', 'Company ID', 'Company', 'Visit Date', 'Notes', 
      'Outcome', 'Stage', 'Status', 'Next Visit Date', 'Days Since Last Visit', 
      'Next Visit Countdown', 'Outcome Category', 'Follow Up Action', 'Owner', 
      'Prospects Match', 'Contact Type', 'Email Sent', 'Competitor'
    ],
    PROSPECTS: [
      'Company ID', 'Address', 'City', 'Zip Code', 'Company Name', 'Industry', 
      'Latitude', 'Longitude', 'Last Outcome', 'Last Outreach Date', 
      'Days Since Last Contact', 'Next Step Due Countdown', 'Next Steps Due Date', 
      'Contact Status', 'Close Probability', 'Priority Score', 
      'UrgencyBand', 'Urgency Score', 'Totals'
    ],
    // Raw Prospects has same headers as Prospects
    RAW_PROSPECTS: [
      'Company ID', 'Address', 'City', 'Zip Code', 'Company Name', 'Industry', 
      'Latitude', 'Longitude', 'Last Outcome', 'Last Outreach Date', 
      'Days Since Last Contact', 'Next Step Due Countdown', 'Next Steps Due Date', 
      'Contact Status', 'Close Probability', 'Priority Score', 
      'UrgencyBand', 'Urgency Score', 'Totals'
    ],
    // Stale Prospects has same headers as Prospects
    STALE_PROSPECTS: [
      'Company ID', 'Address', 'City', 'Zip Code', 'Company Name', 'Industry', 
      'Latitude', 'Longitude', 'Last Outcome', 'Last Outreach Date', 
      'Days Since Last Contact', 'Next Step Due Countdown', 'Next Steps Due Date', 
      'Contact Status', 'Close Probability', 'Priority Score', 
      'UrgencyBand', 'Urgency Score', 'Totals'
    ],
    ACCOUNTS: [
      'Deployed', 'Timestamp', 'Company Name', 'Contact Name', 'Contact Phone', 
      'Contact Role', 'Site Location', 'Mailing Location', 'Roll-Off Fee', 
      'Handling of Metal', 'Roll Off Container Size', 'Notes', 'Payout Price'
    ],
    CONTACTS: [
      'Name', 'Company', 'Account', 'Role', 'Department', 'Phone Number', 
      'Email', 'Address'
    ],
    // Dashboard is a pivot/summary sheet with variable structure
    DASHBOARD: [
      'Stage Pipeline', 'Count', '', 'category', 'Count', '', 'Pipeline by Status', 'Count'
    ],
    SETTINGS: [
      'Category', 'Key', 'Value_1', 'Value_2', 'Value_3', 'Value_4', 'Description', 'Column 8'
    ],
    // Prices sheet has a complex layout with material categories
    PRICES: [
      '', '', '', '', '', '', '', '', '', '', ''
    ],
    SALES: [
      'Sales ID', 'Date', 'Company Name', 'Company ID', 'Material', 'Weight', 'Price', 'Payment Amount'
    ],
    ACTIVE_CONTAINERS: [
      'Company ID', 'Company Name', 'Location Name', 'Location Address', 'City', 'Zip Code', 
      'Current Deployed Asset(s)', 'Container Size'
    ],
    TRANSACTIONS: [
      'Date', 'Company', 'Material', 'Net Weight', 'Price'
    ],
    // System_Schema is metadata - flexible structure
    SYSTEM_SCHEMA: [
      'Object', 'API_Name', 'Label', 'Type', 'Required', 'Options (Aligned & Validated)', 
      'Default_Value', 'Is_System'
    ]
  },

  DEFAULT_OWNER: 'Kyle Buzbee',
  TIMEZONE: 'America/Chicago',
  DATE_FORMAT: 'MM/dd/yyyy'
};
