/**
 * Workspace Tools - Main Menu & Sidebar Controller
 * Google Workspace Setup & Organization Suite
 *
 * Features:
 * - Gmail Label Creation & Auto-Labeling
 * - Gmail Automation (Digest, Lead Capture)
 * - Drive Folder Structure Setup
 * - Drive Inventory & Organization
 *
 * This is the main entry point that creates the menu
 * and controls the sidebar for all workspace tools.
 */

// ============================================
// Main Menu Setup
// ============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menuTitle = typeof COMPANY_CONFIG !== 'undefined' ? COMPANY_CONFIG.MENU_TITLE : '🏢 Workspace Tools';

  ui.createMenu(menuTitle)
    // Sidebar
    .addItem('📱 Open Tools Sidebar', 'showSidebar')
    .addSeparator()

    // Drive Tools
    .addSubMenu(ui.createMenu('📁 Drive Tools')
      .addItem('Create Folder Structure', 'createFolderStructure')
      .addItem('Scan My Drive', 'scanMyDrive')
      .addItem('Scan Specific Folder', 'scanSpecificFolder')
      .addSeparator()
      .addItem('Find Duplicates', 'findDuplicates')
      .addItem('Find Large Files', 'findLargeFiles')
      .addItem('Find Old Files', 'findOldFiles')
      .addItem('Generate Move Recommendations', 'generateRecommendations'))

    // Gmail Tools
    .addSubMenu(ui.createMenu('📧 Gmail Tools')
      .addItem('Create Gmail Labels', 'createAllLabels')
      .addItem('Delete All Created Labels', 'deleteAllCreatedLabels')
      .addItem('List Existing Labels', 'listExistingLabels')
      .addSeparator()
      .addSubMenu(ui.createMenu('🔄 Label Migration')
        .addItem('Scan Labels', 'scanExistingLabels')
        .addItem('Generate Suggestions', 'generateLabelSuggestions')
        .addItem('Preview Migration', 'previewLabelMigration')
        .addItem('Execute Migration', 'executeLabelMigration')
        .addSeparator()
        .addItem('Initialize Migration Sheet', 'initializeLabelMigrationSheet'))
      .addSeparator()
      .addItem('Run Auto-Label on Inbox', 'autoLabelInbox')
      .addItem('Generate Daily Digest', 'generateDailyDigest')
      .addItem('Capture Leads from Email', 'captureLeadsFromEmail')
      .addSeparator()
      .addItem('Archive Old Newsletters', 'archiveOldNewsletters')
      .addItem('Archive Old Promotions', 'archiveOldPromotions'))

    .addSeparator()

    // Reports
    .addSubMenu(ui.createMenu('📊 Reports')
      .addItem('Email Statistics', 'showEmailStatistics')
      .addItem('Unread Summary', 'showUnreadSummary'))

    .addSeparator()

    // Setup
    .addSubMenu(ui.createMenu('⚙️ Setup & Initialize')
      .addItem('Initialize All Sheets', 'initializeAllWorkspaceTools')
      .addSeparator()
      .addItem('Initialize Drive Inventory', 'initializeInventorySheets')
      .addItem('Initialize Gmail Labels', 'initializeLabelSheet')
      .addItem('Initialize Gmail Automation', 'initializeGmailSheets')
      .addItem('Initialize Label Migration', 'initializeLabelMigrationSheet')
      .addItem('Initialize Function Runner', 'initializeFunctionRunnerSheet')
      .addSeparator()
      .addItem('Create All Triggers', 'createAllTriggers')
      .addItem('Remove All Triggers', 'removeAllTriggers')
      .addSeparator()
      .addItem('Setup Function Runner Trigger', 'setupFunctionRunnerTrigger')
      .addItem('Remove Function Runner Trigger', 'removeFunctionRunnerTrigger'))

    .addSeparator()
    .addItem('📊 Open Dashboard', 'showDashboard')
    .addItem('🌐 Open Dashboard in Browser', 'openDashboardInBrowser')

    // Help & About submenu
    .addSubMenu(ui.createMenu('❓ Help & About')
      .addItem('📖 User Manual', 'showUserManual')
      .addItem('📋 Quick Help', 'showHelp')
      .addSeparator()
      .addItem('ℹ️ About This Toolkit', 'showAbout'))

    .addToUi();
}

function doGet(e) {
  try {
    const spreadsheetId = e && e.parameter ? e.parameter.sid : '';
    const ss = getWorkspaceSpreadsheet_(spreadsheetId);
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.spreadsheetId = ss.getId();
    return template.evaluate().setTitle('Workspace Tools Dashboard');
  } catch (error) {
    return HtmlService.createHtmlOutput('<h3>Workspace Tools Web App Setup Required</h3><p>' + error.message + '</p>');
  }
}

function openDashboardInBrowser() {
  const ss = getWorkspaceSpreadsheet_();
  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert('Deploy as Web App first, then retry.');
    return;
  }
  const fullUrl = webAppUrl + '?sid=' + encodeURIComponent(ss.getId());
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial; padding: 12px;">' +
    '<p><strong>Open browser dashboard:</strong></p>' +
    '<p><a href="' + fullUrl + '" target="_blank">' + fullUrl + '</a></p>' +
    '</div>'
  ).setWidth(700).setHeight(160);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Browser Link');
}

function getWorkspaceSpreadsheet_(spreadsheetId) {
  if (spreadsheetId) {
    PropertiesService.getScriptProperties().setProperty('WORKSPACE_SPREADSHEET_ID', spreadsheetId);
    return SpreadsheetApp.openById(spreadsheetId);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('WORKSPACE_SPREADSHEET_ID', active.getId());
    return active;
  }
  const savedId = PropertiesService.getScriptProperties().getProperty('WORKSPACE_SPREADSHEET_ID');
  if (savedId) return SpreadsheetApp.openById(savedId);
  throw new Error('No spreadsheet context found.');
}

// ============================================
// Sidebar Functions
// ============================================
function showSidebar() {
  const title = typeof COMPANY_CONFIG !== 'undefined' ? COMPANY_CONFIG.MENU_TITLE : '🏢 Workspace Tools';
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle(title)
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Show comprehensive help dialog with all features documented
 */
function showHelp() {
  const html = HtmlService.createHtmlOutputFromFile('Help')
    .setTitle('Workspace Tools - Help & Instructions')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Help & Instructions');
}

/**
 * Show user manual with industry-specific documentation
 */
function showUserManual() {
  const title = typeof COMPANY_CONFIG !== 'undefined' ?
    COMPANY_CONFIG.SHORT_NAME + ' User Manual' : 'User Manual';
  const html = HtmlService.createHtmlOutputFromFile('UserManual')
    .setTitle(title)
    .setWidth(700)
    .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/**
 * Show about dialog with toolkit information
 */
function showAbout() {
  const config = typeof COMPANY_CONFIG !== 'undefined' ? COMPANY_CONFIG : {
    NAME: 'Operations Toolkit',
    MENU_ICON: '🏢',
    SHORT_NAME: 'Workspace Tools'
  };

  const industry = typeof LABEL_CONFIG !== 'undefined' && LABEL_CONFIG.INDUSTRY ?
    LABEL_CONFIG.INDUSTRY : 'general';

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; text-align: center; }
      .icon { font-size: 48px; margin-bottom: 10px; }
      h2 { color: #1a73e8; margin-bottom: 5px; }
      .subtitle { color: #5f6368; margin-bottom: 20px; }
      .info { text-align: left; background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 15px 0; }
      .info-row { display: flex; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid #e0e0e0; }
      .info-row:last-child { border-bottom: none; }
      .label { color: #5f6368; }
      .value { font-weight: 500; }
      .footer { margin-top: 20px; font-size: 12px; color: #9aa0a6; }
      .btn { background: #1a73e8; color: white; border: none; padding: 10px 24px; border-radius: 4px; cursor: pointer; margin-top: 15px; }
      .btn:hover { background: #1557b0; }
    </style>
    <div class="icon">${config.MENU_ICON}</div>
    <h2>${config.SHORT_NAME} Operations Toolkit</h2>
    <p class="subtitle">Workspace automation for ${config.NAME}</p>

    <div class="info">
      <div class="info-row">
        <span class="label">Version</span>
        <span class="value">1.0</span>
      </div>
      <div class="info-row">
        <span class="label">Industry</span>
        <span class="value">${industry}</span>
      </div>
      <div class="info-row">
        <span class="label">Support</span>
        <span class="value">jacob@truenorthstrategyops.com</span>
      </div>
    </div>

    <p class="footer">
      Developed by True North Data Strategies<br>
      © ${new Date().getFullYear()} All rights reserved
    </p>

    <button class="btn" onclick="google.script.host.close()">Close</button>
  `)
  .setWidth(350)
  .setHeight(380);

  SpreadsheetApp.getUi().showModalDialog(html, 'About');
}

// ============================================
// Master Initialize
// ============================================
function initializeAllWorkspaceTools() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Initialize Workspace Tools',
    'This will create all sheets for:\n\n' +
    '• Drive Inventory\n' +
    '• Gmail Labels\n' +
    '• Gmail Automation\n' +
    '• Label Migration\n' +
    '• Function Runner\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  // Initialize each module
  try {
    if (typeof initializeInventorySheets === 'function') initializeInventorySheets();
  } catch(e) { Logger.log('Drive Inventory init skipped: ' + e.message); }

  try {
    if (typeof initializeLabelSheet === 'function') initializeLabelSheet();
  } catch(e) { Logger.log('Gmail Labels init skipped: ' + e.message); }

  try {
    if (typeof initializeGmailSheets === 'function') initializeGmailSheets();
  } catch(e) { Logger.log('Gmail Automation init skipped: ' + e.message); }

  try {
    if (typeof initializeLabelMigrationSheet === 'function') initializeLabelMigrationSheet();
  } catch(e) { Logger.log('Label Migration init skipped: ' + e.message); }

  try {
    if (typeof initializeFunctionRunnerSheet === 'function') initializeFunctionRunnerSheet();
  } catch(e) { Logger.log('Function Runner init skipped: ' + e.message); }

  ui.alert('✅ All workspace tools initialized!\n\nRefresh the page to see all sheets.');
}

// ============================================
// Trigger Management
// ============================================
function createAllTriggers() {
  const ui = SpreadsheetApp.getUi();

  // Remove existing first
  removeAllTriggers();

  // Daily digest at 7 AM
  ScriptApp.newTrigger('sendDailyDigestAuto')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  // Lead capture every 2 hours
  ScriptApp.newTrigger('captureLeadsAuto')
    .timeBased()
    .everyHours(2)
    .create();

  // Auto-label every 15 minutes
  ScriptApp.newTrigger('autoLabelInboxSilent')
    .timeBased()
    .everyMinutes(15)
    .create();

  ui.alert('✅ All triggers created!\n\n' +
    '• Daily Digest: 7 AM\n' +
    '• Lead Capture: Every 2 hours\n' +
    '• Auto-Label: Every 15 minutes');
}

function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  Logger.log('All triggers removed');
}

// ============================================
// Quick Actions (called from sidebar)
// ============================================
function quickCreateFolders() {
  createFolderStructure();
}

function quickScanDrive() {
  scanMyDrive();
}

function quickCreateLabels() {
  createAllLabels();
}

function quickDailyDigest() {
  generateDailyDigest();
}

function getToolsSummary() {
  const ss = getWorkspaceSpreadsheet_();
  const sheets = ss.getSheets().map(s => s.getName());

  return {
    sheetCount: sheets.length,
    sheets: sheets,
    hasInventory: sheets.includes('Drive Inventory'),
    hasLabels: sheets.includes('Gmail Labels'),
    hasLabelRules: sheets.includes('Auto-Label Rules'),
    hasCapturedLeads: sheets.includes('Captured Leads'),
    hasEmailStats: sheets.includes('Email Statistics')
  };
}

/**
 * Get company configuration for sidebar
 */
function getCompanyConfig() {
  if (typeof COMPANY_CONFIG !== 'undefined') {
    return {
      name: COMPANY_CONFIG.NAME,
      shortName: COMPANY_CONFIG.SHORT_NAME,
      icon: COMPANY_CONFIG.MENU_ICON,
      email: COMPANY_CONFIG.EMAIL,
      phone: COMPANY_CONFIG.PHONE
    };
  }
  return {
    name: 'Workspace Tools',
    shortName: 'Workspace',
    icon: '🏢',
    email: '',
    phone: ''
  };
}

// ============================================
// Dashboard Functions
// ============================================

/**
 * Show the dashboard
 */
function showDashboard() {
  const menuTitle = typeof COMPANY_CONFIG !== 'undefined' ? COMPANY_CONFIG.MENU_TITLE : 'Operations Toolkit';
  const ss = getWorkspaceSpreadsheet_();
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.spreadsheetId = ss.getId();
  const html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, menuTitle + ' Dashboard');
}

/**
 * Get all dashboard data
 * @return {Object} Dashboard data object
 */
function getDashboardData(spreadsheetId) {
  const ss = getWorkspaceSpreadsheet_(spreadsheetId);

  // Get drive stats
  const driveStats = getDriveStats(ss);

  // Get Gmail stats
  const gmailStats = getGmailStats(ss);

  // Get recent leads
  const recentLeads = getRecentLeads(ss);

  // Get large files
  const largeFiles = getLargeFiles(ss);

  // Get system health
  const health = getOperationsHealth(ss);

  // Get chart data
  const charts = getOperationsChartData(ss);

  const companyName = typeof COMPANY_CONFIG !== 'undefined' ? COMPANY_CONFIG.NAME : 'Operations';

  return {
    companyName: companyName + ' Operations Toolkit',
    lastUpdated: new Date().toLocaleString(),
    driveStats: driveStats,
    gmailStats: gmailStats,
    recentLeads: recentLeads,
    largeFiles: largeFiles,
    health: health,
    charts: charts
  };
}

/**
 * Get Drive inventory statistics
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Object} Drive stats object
 */
function getDriveStats(ss) {
  const stats = {
    totalFiles: 0,
    totalFolders: 0,
    totalSize: '0 MB',
    duplicates: 0,
    recommendations: 0
  };

  // Get inventory sheet
  const invSheet = ss.getSheetByName('Drive Inventory');
  if (invSheet && invSheet.getLastRow() > 1) {
    stats.totalFiles = invSheet.getLastRow() - 1;

    // Calculate total size
    const sizeData = invSheet.getRange(2, 4, stats.totalFiles, 1).getValues();
    let totalKB = 0;
    sizeData.forEach(row => {
      totalKB += row[0] || 0;
    });
    const totalMB = Math.round(totalKB / 1024);
    stats.totalSize = totalMB >= 1024 ? Math.round(totalMB / 1024) + ' GB' : totalMB + ' MB';
  }

  // Get folder count
  const foldersSheet = ss.getSheetByName('Folder Structure');
  if (foldersSheet && foldersSheet.getLastRow() > 1) {
    stats.totalFolders = foldersSheet.getLastRow() - 1;
  }

  // Get duplicates count
  const dupSheet = ss.getSheetByName('Potential Duplicates');
  if (dupSheet && dupSheet.getLastRow() > 1) {
    stats.duplicates = dupSheet.getLastRow() - 1;
  }

  // Get recommendations count
  const recSheet = ss.getSheetByName('Move Recommendations');
  if (recSheet && recSheet.getLastRow() > 1) {
    stats.recommendations = recSheet.getLastRow() - 1;
  }

  return stats;
}

/**
 * Get Gmail automation statistics
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Object} Gmail stats object
 */
function getGmailStats(ss) {
  const stats = {
    labels: 0,
    capturedLeads: 0,
    newLeads: 0,
    autoLabelRules: 0
  };

  // Get labels count
  const labelsSheet = ss.getSheetByName('Gmail Labels');
  if (labelsSheet && labelsSheet.getLastRow() > 1) {
    stats.labels = labelsSheet.getLastRow() - 1;
  }

  // Get captured leads
  const leadsSheet = ss.getSheetByName('Captured Leads');
  if (leadsSheet && leadsSheet.getLastRow() > 1) {
    stats.capturedLeads = leadsSheet.getLastRow() - 1;

    // Count new leads
    const statusData = leadsSheet.getRange(2, 10, stats.capturedLeads, 1).getValues();
    statusData.forEach(row => {
      if (row[0] === 'New') stats.newLeads++;
    });
  }

  // Get auto-label rules
  const rulesSheet = ss.getSheetByName('Auto-Label Rules');
  if (rulesSheet && rulesSheet.getLastRow() > 1) {
    stats.autoLabelRules = rulesSheet.getLastRow() - 1;
  }

  return stats;
}

/**
 * Get recent leads for display
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Array} Array of lead objects
 */
function getRecentLeads(ss) {
  const sheet = ss.getSheetByName('Captured Leads');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const numRows = Math.min(sheet.getLastRow() - 1, 10);
  const data = sheet.getRange(2, 1, numRows, 11).getValues();

  return data.map(row => ({
    date: row[0] ? new Date(row[0]).toLocaleDateString() : '',
    subject: (row[2] || '').substring(0, 50),
    source: row[3] || 'Unknown',
    status: row[9] || 'New'
  })).reverse();
}

/**
 * Get large files for display
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Array} Array of file objects
 */
function getLargeFiles(ss) {
  const sheet = ss.getSheetByName('Drive Inventory');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const largeFiles = [];

  data.forEach(row => {
    const sizeKB = row[3] || 0;
    if (sizeKB > 10240) { // > 10MB
      largeFiles.push({
        name: (row[0] || '').substring(0, 40),
        size: Math.round(sizeKB / 1024) + ' MB',
        path: (row[7] || '').substring(0, 30)
      });
    }
  });

  // Sort by size descending and take top 10
  largeFiles.sort((a, b) => parseInt(b.size) - parseInt(a.size));
  return largeFiles.slice(0, 10);
}

/**
 * Get system health status
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Object} Health status object
 */
function getOperationsHealth(ss) {
  const health = {
    inventory: { status: 'good', message: 'Operational' },
    gmail: { status: 'good', message: 'Connected' },
    labels: { status: 'good', message: 'Operational' },
    triggers: { status: 'good', message: 'Active' }
  };

  // Check Drive Inventory
  const invSheet = ss.getSheetByName('Drive Inventory');
  if (!invSheet) {
    health.inventory = { status: 'warning', message: 'Not initialized' };
  } else if (invSheet.getLastRow() < 2) {
    health.inventory = { status: 'warning', message: 'No data - run scan' };
  }

  // Check Gmail access
  try {
    GmailApp.getAliases();
    health.gmail = { status: 'good', message: 'Connected' };
  } catch (e) {
    health.gmail = { status: 'error', message: 'Not authorized' };
  }

  // Check Labels sheet
  const labelsSheet = ss.getSheetByName('Gmail Labels');
  if (!labelsSheet) {
    health.labels = { status: 'warning', message: 'Not initialized' };
  }

  // Check Triggers
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    health.triggers = { status: 'warning', message: 'No triggers set' };
  } else {
    health.triggers = { status: 'good', message: triggers.length + ' active' };
  }

  return health;
}

/**
 * Get chart data for dashboard
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Object} Chart data object
 */
function getOperationsChartData(ss) {
  const charts = {
    categoryLabels: [],
    categoryCounts: [],
    suggestedLabels: [],
    suggestedCounts: [],
    leadSourceLabels: [],
    leadSourceCounts: []
  };

  // Get file categories from inventory
  const invSheet = ss.getSheetByName('Drive Inventory');
  if (invSheet && invSheet.getLastRow() > 1) {
    const categoryData = invSheet.getRange(2, 11, invSheet.getLastRow() - 1, 1).getValues();
    const categoryMap = {};

    categoryData.forEach(row => {
      const cat = row[0] || 'Other';
      categoryMap[cat] = (categoryMap[cat] || 0) + 1;
    });

    const sortedCategories = Object.entries(categoryMap).sort((a, b) => b[1] - a[1]).slice(0, 6);
    charts.categoryLabels = sortedCategories.map(c => c[0]);
    charts.categoryCounts = sortedCategories.map(c => c[1]);

    // Get suggested folders
    const suggestedData = invSheet.getRange(2, 12, invSheet.getLastRow() - 1, 1).getValues();
    const suggestedMap = {};

    suggestedData.forEach(row => {
      const folder = row[0] || 'Review Needed';
      suggestedMap[folder] = (suggestedMap[folder] || 0) + 1;
    });

    const sortedSuggested = Object.entries(suggestedMap).sort((a, b) => b[1] - a[1]).slice(0, 5);
    charts.suggestedLabels = sortedSuggested.map(s => s[0]);
    charts.suggestedCounts = sortedSuggested.map(s => s[1]);
  }

  // Get lead sources
  const leadsSheet = ss.getSheetByName('Captured Leads');
  if (leadsSheet && leadsSheet.getLastRow() > 1) {
    const sourceData = leadsSheet.getRange(2, 4, leadsSheet.getLastRow() - 1, 1).getValues();
    const sourceMap = {};

    sourceData.forEach(row => {
      const source = row[0] || 'Unknown';
      sourceMap[source] = (sourceMap[source] || 0) + 1;
    });

    const sortedSources = Object.entries(sourceMap).sort((a, b) => b[1] - a[1]).slice(0, 5);
    charts.leadSourceLabels = sortedSources.map(s => s[0]);
    charts.leadSourceCounts = sortedSources.map(s => s[1]);
  }

  // Ensure there's always some data for charts
  if (charts.categoryLabels.length === 0) {
    charts.categoryLabels = ['No data'];
    charts.categoryCounts = [0];
  }
  if (charts.suggestedLabels.length === 0) {
    charts.suggestedLabels = ['No data'];
    charts.suggestedCounts = [0];
  }
  if (charts.leadSourceLabels.length === 0) {
    charts.leadSourceLabels = ['No data'];
    charts.leadSourceCounts = [0];
  }

  return charts;
}

