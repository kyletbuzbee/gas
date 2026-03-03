# K&L Recycling CRM - Agent Guidelines

## Project Overview
Google Apps Script CRM for K&L Recycling. Uses clasp for deployment to Google Apps Script.

## Build & Deploy Commands

```bash
# Deploy to Google Apps Script
clasp push

# Deploy as web app
clasp deploy

# Open in Apps Script editor
clasp open

# View logs
clasp logs

# Pull remote changes
clasp pull
```

## Code Style Guidelines

### Language & Syntax
- **Language**: Google Apps Script (JavaScript ES6+ with V8 runtime)
- **Indentation**: 2 spaces
- **Quotes**: Single quotes for strings
- **Semicolons**: Required
- **Line length**: 100 characters max

### Naming Conventions
```javascript
// Constants - UPPER_SNAKE_CASE
const CONFIG = { ... };
const DEFAULT_OWNER = 'Kyle Buzbee';

// Functions - camelCase
function getCompanyProfile(companyId) { ... }
function renderDashboard() { ... }

// Variables - camelCase
const companyName = record['Company Name'];
let totalRevenue = 0;

// Sheet keys in CONFIG - SCREAMING_SNAKE_CASE
CONFIG.SHEETS.OUTREACH
CONFIG.SHEETS.ACTIVE_CONTAINERS

// Private helpers - camelCase with underscore prefix (if needed)
function _sanitizeData(data) { ... }
```

### Documentation (JSDoc Required)
```javascript
/**
 * Brief description of what the function does
 * @param {string} paramName - Parameter description
 * @param {Array<Object>} records - Array of record objects
 * @returns {Object} - Description of return value
 * @throws {Error} - When/why error is thrown
 */
function myFunction(paramName, records) {
  // Implementation
}
```

### Error Handling Pattern
```javascript
function safeOperation() {
  try {
    // Operation that might fail
    const result = riskyCall();
    return { success: true, data: result };
  } catch (error) {
    Logger.log('Error in safeOperation: ' + error.message);
    return { success: false, error: error.message };
  }
}

// For write operations - always use LockService
function writeOperation() {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      throw new Error('Could not acquire lock');
    }
    // Write operations here
    return { success: true };
  } catch (error) {
    Logger.log('Error: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}
```

### Performance Patterns
```javascript
// ✅ GOOD: Batch fetch with getMultipleRecords
const allData = DataLayer.getMultipleRecords([
  'PROSPECTS', 'ACTIVE_CONTAINERS', 'OUTREACH'
]);

// ❌ BAD: Sequential individual calls
const prospects = DataLayer.getRecords('PROSPECTS');
const active = DataLayer.getRecords('ACTIVE_CONTAINERS'); // N+1 problem

// ✅ GOOD: Process in chunks with delays for large datasets
const batchSize = 10;
for (let i = 0; i < items.length; i += batchSize) {
  const batch = items.slice(i, i + batchSize);
  processBatch(batch);
  if (i % 50 === 0) Utilities.sleep(100);
}
```

### Sheet Data Access
```javascript
// Always use CONFIG constants for sheet names
const sheet = ss.getSheetByName(CONFIG.SHEETS.OUTREACH);

// Column indices are 0-based when using getValues()
const companyId = row[0];  // Column A
const companyName = row[4]; // Column E

// Check for empty/undefined before processing
if (!companyId || !companyName) return;
```

### Key Architectural Patterns

1. **DataLayer Pattern**: All sheet operations go through `DataLayer` object
2. **CONFIG Centralization**: Sheet names and headers defined in `Config.gs`
3. **LockService**: Required for all write operations to prevent data corruption
4. **Sanitization**: Use `sanitizeForFrontend()` before sending Dates to HTML
5. **Filtering Won Accounts**: Check `getWonCompanyIds()` when displaying prospects

### File Organization
- `Config.gs` - Centralized constants
- `DataLayer.gs` - Database abstraction layer
- `CRMService.gs` - Business logic
- `WebApp.gs` - Web app entry points
- `SidebarScript.gs` - Google Sheets sidebar functions
- `Automations.gs` - Triggers and automation
- `*.html` - Frontend templates

### Testing Approach
No automated test runner. Test by:
1. Running functions in Apps Script editor
2. Checking `View > Logs` (or `clasp logs`)
3. Testing web app in browser
4. Verifying sheet data changes

### Common Gotchas
- Dates from sheets are Date objects - format before sending to frontend
- Sheet headers must match CONFIG.HEADERS exactly
- Company ID is the single source of truth for relationships
- Revenue always comes from SALES sheet (not Transactions)
- Won accounts should be filtered from prospect views
