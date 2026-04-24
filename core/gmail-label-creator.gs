/**
 * Gmail Label Creator & Organizer
 * TNDS Workspace Automation Kit
 *
 * Features:
 * - Create label hierarchy for {{INDUSTRY_NAME}} business
 * - Auto-label incoming emails based on rules
 * - Color-coded labels
 * - Filter creation for common email types
 */

// ============================================
// Configuration - {{INDUSTRY_NAME}} Labels
// ============================================
const LABEL_CONFIG = {
  // Sorting prefix - add prefix to push labels to bottom (e.g., "zz-") or top (e.g., "00-")
  // Leave empty for no prefix (alphabetical with other labels)
  LABEL_PREFIX: '',

  // Industry type for reference
  INDUSTRY: '{{INDUSTRY_KEY}}',

  // Label structure with colors (Gmail color names)
  LABELS: {{LABEL_STRUCTURE}},

  // Auto-labeling rules based on email content
  AUTO_LABEL_RULES: {{AUTO_LABEL_RULES}}
};

// ============================================
// Menu Setup (disabled - using unified menu in main Tools file)
// ============================================
function _onOpenGmailLabels() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Gmail Labels')
    .addItem('🏷️ Create All Labels', 'createAllLabels')
    .addItem('📋 List Existing Labels', 'listExistingLabels')
    .addSeparator()
    .addSubMenu(ui.createMenu('🤖 Auto-Labeling')
      .addItem('Run Auto-Label on Inbox', 'autoLabelInbox')
      .addItem('Create Auto-Label Trigger', 'createAutoLabelTrigger')
      .addItem('Remove Auto-Label Trigger', 'removeAutoLabelTrigger'))
    .addSeparator()
    .addItem('⚙️ Initialize Sheet', 'initializeLabelSheet')
    .addToUi();
}

// ============================================
// Sheet Initialization
// ============================================
function initializeLabelSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Gmail Labels');
  if (!sheet) {
    sheet = ss.insertSheet('Gmail Labels');
  }
  sheet.clear();

  const headers = ['Label Name', 'Status', 'Email Count', 'Created'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 100);
  sheet.setFrozenRows(1);

  // Also create rules sheet
  let rulesSheet = ss.getSheetByName('Auto-Label Rules');
  if (!rulesSheet) {
    rulesSheet = ss.insertSheet('Auto-Label Rules');
  }
  rulesSheet.clear();

  const ruleHeaders = ['Label', 'From Contains', 'Subject Contains', 'Active'];
  rulesSheet.getRange(1, 1, 1, ruleHeaders.length).setValues([ruleHeaders]);
  rulesSheet.getRange(1, 1, 1, ruleHeaders.length)
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Add default rules
  const rules = LABEL_CONFIG.AUTO_LABEL_RULES.map(r => [
    r.label,
    r.from.join(', '),
    r.subject.join(', '),
    true
  ]);

  if (rules.length > 0) {
    rulesSheet.getRange(2, 1, rules.length, 4).setValues(rules);
  }

  // Checkbox for Active
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  rulesSheet.getRange(2, 4, 50, 1).setDataValidation(checkboxRule);

  rulesSheet.setColumnWidth(1, 200);
  rulesSheet.setColumnWidth(2, 200);
  rulesSheet.setColumnWidth(3, 200);
  rulesSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Label sheets initialized!');
}

// ============================================
// Label Creation
// ============================================
function createAllLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Gmail Labels');

  if (!sheet) {
    initializeLabelSheet();
    sheet = ss.getSheetByName('Gmail Labels');
  }

  const results = [];
  let created = 0;
  let existing = 0;

  // Get prefix for sorting control
  const prefix = LABEL_CONFIG.LABEL_PREFIX || '';

  // Get existing labels
  const existingLabels = GmailApp.getUserLabels().map(l => l.getName());

  // Create parent labels and children
  for (const [parentName, config] of Object.entries(LABEL_CONFIG.LABELS)) {
    // Apply prefix for sorting
    const fullParentName = prefix + parentName;

    // Create parent label
    if (!existingLabels.includes(fullParentName)) {
      GmailApp.createLabel(fullParentName);
      results.push([fullParentName, 'Created', 0, new Date()]);
      created++;
    } else {
      const label = GmailApp.getUserLabelByName(fullParentName);
      results.push([fullParentName, 'Exists', label.getThreads(0, 1).length > 0 ? label.getThreads().length : 0, '']);
      existing++;
    }

    // Create children
    if (config.children) {
      for (const childName of Object.keys(config.children)) {
        // Apply prefix to child (child already has parent path like "Clients/Buyers")
        const fullChildName = prefix + childName;

        if (!existingLabels.includes(fullChildName)) {
          GmailApp.createLabel(fullChildName);
          results.push([fullChildName, 'Created', 0, new Date()]);
          created++;
        } else {
          const label = GmailApp.getUserLabelByName(fullChildName);
          results.push([fullChildName, 'Exists', '', '']);
          existing++;
        }
      }
    }
  }

  // Write to sheet
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clear();
  }

  if (results.length > 0) {
    sheet.getRange(2, 1, results.length, 4).setValues(results);
  }

  // Color code status
  const statusRange = sheet.getRange(2, 2, results.length, 1);
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Created')
    .setBackground('#c8e6c9')
    .setRanges([statusRange])
    .build();
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Exists')
    .setBackground('#e3f2fd')
    .setRanges([statusRange])
    .build();
  sheet.setConditionalFormatRules([rule1, rule2]);

  SpreadsheetApp.getUi().alert(`✅ Label creation complete!\n\nCreated: ${created}\nAlready existed: ${existing}`);
}

function listExistingLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Gmail Labels');

  if (!sheet) {
    initializeLabelSheet();
    sheet = ss.getSheetByName('Gmail Labels');
  }

  const labels = GmailApp.getUserLabels();
  const results = [];

  labels.forEach(label => {
    const threads = label.getThreads(0, 100);
    results.push([
      label.getName(),
      'Exists',
      threads.length,
      ''
    ]);
  });

  // Sort alphabetically
  results.sort((a, b) => a[0].localeCompare(b[0]));

  // Clear and write
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clear();
  }

  if (results.length > 0) {
    sheet.getRange(2, 1, results.length, 4).setValues(results);
  }

  SpreadsheetApp.getUi().alert(`Found ${results.length} labels in Gmail.`);
}

// ============================================
// Auto-Labeling
// ============================================
function autoLabelInbox() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName('Auto-Label Rules');

  // Get rules from sheet or use defaults
  let rules = LABEL_CONFIG.AUTO_LABEL_RULES;

  if (rulesSheet && rulesSheet.getLastRow() > 1) {
    const data = rulesSheet.getDataRange().getValues();
    rules = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === true) {  // Active checkbox
        rules.push({
          label: data[i][0],
          from: data[i][1] ? data[i][1].split(',').map(s => s.trim()) : [],
          subject: data[i][2] ? data[i][2].split(',').map(s => s.trim()) : []
        });
      }
    }
  }

  // Get unread threads from inbox
  const threads = GmailApp.search('is:unread in:inbox', 0, 50);
  let labeled = 0;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const firstMessage = messages[0];
    const from = firstMessage.getFrom().toLowerCase();
    const subject = firstMessage.getSubject().toLowerCase();

    for (const rule of rules) {
      let matched = false;

      // Check from
      for (const fromPattern of rule.from) {
        if (from.includes(fromPattern.toLowerCase())) {
          matched = true;
          break;
        }
      }

      // Check subject
      if (!matched) {
        for (const subjectPattern of rule.subject) {
          if (subject.includes(subjectPattern.toLowerCase())) {
            matched = true;
            break;
          }
        }
      }

      if (matched) {
        // Apply prefix to label name
        const prefix = LABEL_CONFIG.LABEL_PREFIX || '';
        const fullLabelName = prefix + rule.label;
        const label = GmailApp.getUserLabelByName(fullLabelName);
        if (label) {
          thread.addLabel(label);
          labeled++;
          Logger.log('Labeled: ' + subject + ' -> ' + fullLabelName);
        } else {
          // Create label if it doesn't exist
          GmailApp.createLabel(fullLabelName);
          const newLabel = GmailApp.getUserLabelByName(fullLabelName);
          thread.addLabel(newLabel);
          labeled++;
        }
        break;  // Only apply first matching rule
      }
    }
  });

  SpreadsheetApp.getUi().alert(`✅ Auto-labeling complete!\n\nProcessed: ${threads.length} threads\nLabeled: ${labeled}`);
}

function createAutoLabelTrigger() {
  // Remove existing
  removeAutoLabelTrigger();

  // Create trigger to run every 15 minutes
  ScriptApp.newTrigger('autoLabelInboxSilent')
    .timeBased()
    .everyMinutes(15)
    .create();

  SpreadsheetApp.getUi().alert('✅ Auto-labeling scheduled to run every 15 minutes.');
}

function removeAutoLabelTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoLabelInboxSilent') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// Silent version for automated runs
function autoLabelInboxSilent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rulesSheet = ss.getSheetByName('Auto-Label Rules');

  let rules = LABEL_CONFIG.AUTO_LABEL_RULES;

  if (rulesSheet && rulesSheet.getLastRow() > 1) {
    const data = rulesSheet.getDataRange().getValues();
    rules = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === true) {
        rules.push({
          label: data[i][0],
          from: data[i][1] ? data[i][1].split(',').map(s => s.trim()) : [],
          subject: data[i][2] ? data[i][2].split(',').map(s => s.trim()) : []
        });
      }
    }
  }

  const threads = GmailApp.search('is:unread in:inbox', 0, 50);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const firstMessage = messages[0];
    const from = firstMessage.getFrom().toLowerCase();
    const subject = firstMessage.getSubject().toLowerCase();

    for (const rule of rules) {
      let matched = false;

      for (const fromPattern of rule.from) {
        if (from.includes(fromPattern.toLowerCase())) {
          matched = true;
          break;
        }
      }

      if (!matched) {
        for (const subjectPattern of rule.subject) {
          if (subject.includes(subjectPattern.toLowerCase())) {
            matched = true;
            break;
          }
        }
      }

      if (matched) {
        // Apply prefix to label name
        const prefix = LABEL_CONFIG.LABEL_PREFIX || '';
        const fullLabelName = prefix + rule.label;
        let label = GmailApp.getUserLabelByName(fullLabelName);
        if (!label) {
          GmailApp.createLabel(fullLabelName);
          label = GmailApp.getUserLabelByName(fullLabelName);
        }
        thread.addLabel(label);
        Logger.log('Auto-labeled: ' + subject + ' -> ' + fullLabelName);
        break;
      }
    }
  });
}

// ============================================
// Utility Functions
// ============================================

/**
 * Delete a specific label
 */
function deleteLabel(labelName) {
  const label = GmailApp.getUserLabelByName(labelName);
  if (label) {
    label.deleteLabel();
    Logger.log('Deleted label: ' + labelName);
    return true;
  }
  return false;
}

/**
 * Move all emails from one label to another
 */
function moveEmailsBetweenLabels(fromLabelName, toLabelName) {
  const fromLabel = GmailApp.getUserLabelByName(fromLabelName);
  const toLabel = GmailApp.getUserLabelByName(toLabelName);

  if (!fromLabel || !toLabel) {
    Logger.log('One or both labels not found');
    return 0;
  }

  const threads = fromLabel.getThreads();
  let count = 0;

  threads.forEach(thread => {
    thread.addLabel(toLabel);
    thread.removeLabel(fromLabel);
    count++;
  });

  Logger.log(`Moved ${count} threads from ${fromLabelName} to ${toLabelName}`);
  return count;
}

/**
 * Archive all emails with a specific label
 */
function archiveByLabel(labelName) {
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) return 0;

  const threads = label.getThreads();
  threads.forEach(thread => {
    thread.moveToArchive();
  });

  Logger.log(`Archived ${threads.length} threads with label ${labelName}`);
  return threads.length;
}

// ============================================
// Delete/Undo Labels
// ============================================

/**
 * Delete all labels created by this script (with confirmation)
 * This is the UNDO function for createAllLabels
 */
function deleteAllCreatedLabels() {
  const ui = SpreadsheetApp.getUi();

  // Get all label names from config
  const labelsToDelete = [];
  const prefix = LABEL_CONFIG.LABEL_PREFIX || '';

  for (const [parentName, config] of Object.entries(LABEL_CONFIG.LABELS)) {
    const fullParentName = prefix + parentName;
    labelsToDelete.push(fullParentName);

    if (config.children) {
      for (const childName of Object.keys(config.children)) {
        // Child names already include parent path, so add prefix to start
        const fullChildName = prefix + childName;
        labelsToDelete.push(fullChildName);
      }
    }
  }

  // Show confirmation with list
  const response = ui.alert(
    '⚠️ Delete All Created Labels?',
    `This will delete ${labelsToDelete.length} labels:\n\n` +
    labelsToDelete.slice(0, 10).join('\n') +
    (labelsToDelete.length > 10 ? `\n... and ${labelsToDelete.length - 10} more` : '') +
    '\n\nEmails will NOT be deleted, only the labels will be removed.\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled. No labels were deleted.');
    return;
  }

  // Delete labels
  let deleted = 0;
  let notFound = 0;
  const errors = [];

  // Delete children first, then parents (reverse order)
  const reversedLabels = [...labelsToDelete].reverse();

  for (const labelName of reversedLabels) {
    try {
      const label = GmailApp.getUserLabelByName(labelName);
      if (label) {
        label.deleteLabel();
        deleted++;
        Logger.log('Deleted label: ' + labelName);
      } else {
        notFound++;
      }
    } catch (e) {
      errors.push(labelName + ': ' + e.message);
      Logger.log('Error deleting ' + labelName + ': ' + e.message);
    }
  }

  // Update sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Gmail Labels');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clear();
  }

  // Show result
  let message = `✅ Label deletion complete!\n\nDeleted: ${deleted}\nNot found: ${notFound}`;
  if (errors.length > 0) {
    message += `\nErrors: ${errors.length}`;
  }
  ui.alert(message);
}

/**
 * Get prefixed label name
 */
function getPrefixedLabelName(labelName) {
  const prefix = LABEL_CONFIG.LABEL_PREFIX || '';
  return prefix + labelName;
}

// ============================================
// Label Migration System
// ============================================

/**
 * Label migration keyword mappings
 * Maps keywords in existing label names to new structured labels
 */
const LABEL_MIGRATION_MAPPINGS = {
  // Client-related
  'buyer': 'Clients/Buyers',
  'buyers': 'Clients/Buyers',
  'seller': 'Clients/Sellers',
  'sellers': 'Clients/Sellers',
  'prospect': 'Clients/Prospects',
  'prospects': 'Clients/Prospects',
  'past client': 'Clients/Past Clients',
  'former client': 'Clients/Past Clients',

  // Lead sources
  'zillow': 'Leads/Zillow',
  'z-lead': 'Leads/Zillow',
  'realtor.com': 'Leads/Realtor.com',
  'realtor com': 'Leads/Realtor.com',
  'website': 'Leads/Website',
  'web lead': 'Leads/Website',
  'referral': 'Leads/Referral',
  'open house': 'Leads/Open House',
  'openhouse': 'Leads/Open House',

  // Transactions
  'active deal': 'Transactions/Active',
  'under contract': 'Transactions/Pending',
  'pending': 'Transactions/Pending',
  'closed': 'Transactions/Closed',
  'closing': 'Transactions/Closed',

  // Listings
  'active listing': 'Listings/Active',
  'coming soon': 'Listings/Coming Soon',
  'sold': 'Listings/Sold',

  // Vendors
  'title': 'Vendors/Title',
  'escrow': 'Vendors/Title',
  'lender': 'Vendors/Lenders',
  'mortgage': 'Vendors/Lenders',
  'loan': 'Vendors/Lenders',
  'inspector': 'Vendors/Inspectors',
  'inspection': 'Vendors/Inspectors',
  'contractor': 'Vendors/Contractors',
  'repair': 'Vendors/Contractors',

  // Admin
  'brokerage': 'Admin/Brokerage',
  'broker': 'Admin/Brokerage',
  'mls': 'Admin/MLS',
  'license': 'Admin/License',

  // Special
  'action': 'Action Required',
  'urgent': 'Action Required',
  'todo': 'Action Required',
  'waiting': 'Waiting For Response',
  'follow up': 'Waiting For Response',
  'followup': 'Waiting For Response'
};

/**
 * Initialize the Label Migration sheet
 */
function initializeLabelMigrationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName('Label Migration');
  if (!sheet) {
    sheet = ss.insertSheet('Label Migration');
  }
  sheet.clear();

  // Headers
  const headers = ['Current Label', 'Email Count', 'Suggested New Label', 'Move Emails?', 'Status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#9c27b0')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Column widths
  sheet.setColumnWidth(1, 200);  // Current Label
  sheet.setColumnWidth(2, 100);  // Email Count
  sheet.setColumnWidth(3, 200);  // Suggested New Label
  sheet.setColumnWidth(4, 100);  // Move Emails?
  sheet.setColumnWidth(5, 100);  // Status
  sheet.setFrozenRows(1);

  // Create dropdown options for new labels column
  const newLabelOptions = _getAllConfiguredLabels();

  // Add checkbox for Move Emails column
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, 4, 200, 1).setDataValidation(checkboxRule);

  // Add dropdown for Suggested New Label column
  if (newLabelOptions.length > 0) {
    const dropdownRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(newLabelOptions, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 3, 200, 1).setDataValidation(dropdownRule);
  }

  // Add status conditional formatting
  const statusRange = sheet.getRange('E:E');
  const rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Done')
    .setBackground('#c8e6c9')
    .setRanges([statusRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#fff3e0')
    .setRanges([statusRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Skipped')
    .setBackground('#e0e0e0')
    .setRanges([statusRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Error')
    .setBackground('#ffcdd2')
    .setRanges([statusRange])
    .build());
  sheet.setConditionalFormatRules(rules);

  SpreadsheetApp.getUi().alert('✅ Label Migration sheet initialized!\n\nNext: Run "Scan Labels" to populate the sheet.');
}

/**
 * Get all configured label names from LABEL_CONFIG
 */
function _getAllConfiguredLabels() {
  const labels = [];
  const prefix = LABEL_CONFIG.LABEL_PREFIX || '';

  for (const [parentName, config] of Object.entries(LABEL_CONFIG.LABELS)) {
    labels.push(prefix + parentName);
    if (config.children) {
      for (const childName of Object.keys(config.children)) {
        labels.push(prefix + childName);
      }
    }
  }

  return labels.sort();
}

/**
 * Scan existing Gmail labels and populate the migration sheet
 */
function scanExistingLabels() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Label Migration');

  if (!sheet) {
    initializeLabelMigrationSheet();
    sheet = ss.getSheetByName('Label Migration');
  }

  // Clear existing data
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clear();
  }

  const ui = SpreadsheetApp.getUi();
  ui.alert('Scanning labels... This may take a moment.');

  // Get all user labels
  const labels = GmailApp.getUserLabels();
  const results = [];
  const configuredLabels = _getAllConfiguredLabels();

  labels.forEach(label => {
    const labelName = label.getName();

    // Skip labels that are already in our configured structure
    if (configuredLabels.includes(labelName)) {
      return;
    }

    // Get email count (limited query for performance)
    let emailCount = 0;
    try {
      const threads = label.getThreads(0, 500);
      emailCount = threads.length;
      if (emailCount === 500) {
        emailCount = '500+';  // Indicate there may be more
      }
    } catch (e) {
      emailCount = 'Error';
    }

    // Get suggested new label
    const suggestion = _getSuggestedLabel(labelName);

    results.push([
      labelName,
      emailCount,
      suggestion,
      false,  // Move Emails checkbox
      'Pending'
    ]);
  });

  // Sort by email count descending (labels with most emails first)
  results.sort((a, b) => {
    const countA = typeof a[1] === 'number' ? a[1] : 999;
    const countB = typeof b[1] === 'number' ? b[1] : 999;
    return countB - countA;
  });

  // Write results
  if (results.length > 0) {
    sheet.getRange(2, 1, results.length, 5).setValues(results);
  }

  ui.alert(`✅ Scan complete!\n\nFound ${results.length} labels to review.\n\n` +
    'Review the "Suggested New Label" column, adjust as needed, then check "Move Emails?" for labels you want to migrate.');
}

/**
 * Get suggested label based on keyword matching
 */
function _getSuggestedLabel(labelName) {
  const lowerName = labelName.toLowerCase();
  const prefix = LABEL_CONFIG.LABEL_PREFIX || '';

  // Check each mapping
  for (const [keyword, newLabel] of Object.entries(LABEL_MIGRATION_MAPPINGS)) {
    if (lowerName.includes(keyword.toLowerCase())) {
      return prefix + newLabel;
    }
  }

  return '';  // No suggestion
}

/**
 * Generate/refresh label suggestions for existing data
 */
function generateLabelSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Label Migration');

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No labels to process. Run "Scan Labels" first.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const currentLabel = data[i][0];
    const existingSuggestion = data[i][2];

    // Only update if no suggestion exists
    if (!existingSuggestion || existingSuggestion === '') {
      const newSuggestion = _getSuggestedLabel(currentLabel);
      if (newSuggestion) {
        sheet.getRange(i + 1, 3).setValue(newSuggestion);
        updated++;
      }
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Updated ${updated} suggestions.`);
}

/**
 * Preview the label migration before executing
 */
function previewLabelMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Label Migration');

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No labels to migrate. Run "Scan Labels" first.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const migrations = [];
  let totalEmails = 0;

  for (let i = 1; i < data.length; i++) {
    const moveChecked = data[i][3];
    const status = data[i][4];

    if (moveChecked === true && status !== 'Done') {
      const currentLabel = data[i][0];
      const emailCount = data[i][1];
      const newLabel = data[i][2];

      if (!newLabel) {
        migrations.push(`⚠️ ${currentLabel} - NO TARGET LABEL SET`);
      } else {
        migrations.push(`• ${currentLabel} → ${newLabel} (${emailCount} emails)`);
        if (typeof emailCount === 'number') {
          totalEmails += emailCount;
        }
      }
    }
  }

  if (migrations.length === 0) {
    SpreadsheetApp.getUi().alert('No labels selected for migration.\n\nCheck the "Move Emails?" column for labels you want to migrate.');
    return;
  }

  const preview = `LABEL MIGRATION PREVIEW\n\n` +
    `Labels to migrate: ${migrations.length}\n` +
    `Estimated emails: ${totalEmails}+\n\n` +
    migrations.slice(0, 15).join('\n') +
    (migrations.length > 15 ? `\n\n... and ${migrations.length - 15} more` : '') +
    `\n\nRun "Execute Migration" to proceed.`;

  SpreadsheetApp.getUi().alert(preview);
}

/**
 * Execute the label migration for checked rows
 */
function executeLabelMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Label Migration');
  const ui = SpreadsheetApp.getUi();

  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('No labels to migrate. Run "Scan Labels" first.');
    return;
  }

  // Confirm execution
  const response = ui.alert(
    '⚠️ Execute Label Migration',
    'This will move all emails from old labels to new labels.\n\n' +
    'The old labels will NOT be deleted (you can delete them manually after verifying).\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Migration cancelled.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let migrated = 0;
  let skipped = 0;
  let errors = 0;
  const prefix = LABEL_CONFIG.LABEL_PREFIX || '';

  for (let i = 1; i < data.length; i++) {
    const moveChecked = data[i][3];
    const status = data[i][4];

    if (moveChecked === true && status !== 'Done') {
      const currentLabelName = data[i][0];
      const newLabelName = data[i][2];

      if (!newLabelName) {
        sheet.getRange(i + 1, 5).setValue('Skipped');
        skipped++;
        continue;
      }

      try {
        const currentLabel = GmailApp.getUserLabelByName(currentLabelName);
        if (!currentLabel) {
          sheet.getRange(i + 1, 5).setValue('Error: Label not found');
          errors++;
          continue;
        }

        // Get or create new label
        let newLabel = GmailApp.getUserLabelByName(newLabelName);
        if (!newLabel) {
          newLabel = GmailApp.createLabel(newLabelName);
        }

        // Move all emails
        const threads = currentLabel.getThreads();
        let movedCount = 0;

        threads.forEach(thread => {
          thread.addLabel(newLabel);
          thread.removeLabel(currentLabel);
          movedCount++;
        });

        sheet.getRange(i + 1, 5).setValue('Done');
        sheet.getRange(i + 1, 2).setValue(movedCount);  // Update actual count
        migrated++;

        Logger.log(`Migrated ${movedCount} emails from "${currentLabelName}" to "${newLabelName}"`);

      } catch (e) {
        sheet.getRange(i + 1, 5).setValue('Error: ' + e.message);
        errors++;
        Logger.log(`Error migrating ${currentLabelName}: ${e.message}`);
      }
    }
  }

  const summary = `✅ Label Migration Complete!\n\n` +
    `Labels migrated: ${migrated}\n` +
    `Skipped (no target): ${skipped}\n` +
    `Errors: ${errors}\n\n` +
    `Note: Old labels still exist. You can delete them manually after verifying the migration.`;

  ui.alert(summary);
}
