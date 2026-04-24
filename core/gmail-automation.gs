/**
 * Gmail Automation Suite
 * TNDS Workspace Automation Kit
 *
 * Features:
 * - Daily email digest
 * - Lead capture from emails
 * - Auto-responses
 * - Email statistics
 * - Cleanup utilities
 */

// ============================================
// Configuration
// ============================================
const GMAIL_CONFIG = {
  SHEETS: {
    CAPTURED_LEADS: 'Captured Leads',
    EMAIL_STATS: 'Email Statistics',
    TEMPLATES: 'Response Templates'
  },
  // Patterns to identify lead emails
  LEAD_PATTERNS: {
    from: ['zillow.com', 'realtor.com', 'homes.com', 'trulia.com', 'redfin.com'],
    subject: ['new lead', 'inquiry', 'interested in', 'showing request', 'contact request', 'property inquiry']
  },
  // Auto-archive rules (emails older than X days)
  AUTO_ARCHIVE: {
    newsletters: 7,
    promotions: 3,
    social: 7
  }
};

// ============================================
// Menu Setup (disabled - using unified menu in main Tools file)
// ============================================
function _onOpenGmailAutomation() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Gmail Automation')
    .addItem('📬 Generate Daily Digest', 'generateDailyDigest')
    .addItem('🎯 Capture Leads from Email', 'captureLeadsFromEmail')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Reports')
      .addItem('Email Statistics', 'showEmailStatistics')
      .addItem('Unread Summary', 'showUnreadSummary')
      .addItem('Response Time Report', 'showResponseTimeReport'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🧹 Cleanup')
      .addItem('Archive Old Newsletters', 'archiveOldNewsletters')
      .addItem('Archive Old Promotions', 'archiveOldPromotions')
      .addItem('Empty Trash', 'emptyTrash'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('Initialize Sheets', 'initializeGmailSheets')
      .addItem('Create Daily Digest Trigger', 'createDigestTrigger')
      .addItem('Create Lead Capture Trigger', 'createLeadCaptureTrigger'))
    .addToUi();
}

// ============================================
// Sheet Initialization
// ============================================
function initializeGmailSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Captured Leads sheet
  let leadsSheet = ss.getSheetByName(GMAIL_CONFIG.SHEETS.CAPTURED_LEADS);
  if (!leadsSheet) {
    leadsSheet = ss.insertSheet(GMAIL_CONFIG.SHEETS.CAPTURED_LEADS);
  }
  leadsSheet.clear();

  const leadHeaders = ['Date Captured', 'From', 'Subject', 'Source', 'Name', 'Email', 'Phone', 'Property', 'Message Preview', 'Status', 'Email Link'];
  leadsSheet.getRange(1, 1, 1, leadHeaders.length).setValues([leadHeaders]);
  leadsSheet.getRange(1, 1, 1, leadHeaders.length)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  leadsSheet.setFrozenRows(1);
  leadsSheet.setColumnWidth(2, 200);
  leadsSheet.setColumnWidth(3, 300);
  leadsSheet.setColumnWidth(9, 400);

  // Status dropdown
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New', 'Contacted', 'Added to CRM', 'Not a Lead', 'Spam'])
    .build();
  leadsSheet.getRange(2, 10, 500, 1).setDataValidation(statusRule);

  // Email Stats sheet
  let statsSheet = ss.getSheetByName(GMAIL_CONFIG.SHEETS.EMAIL_STATS);
  if (!statsSheet) {
    statsSheet = ss.insertSheet(GMAIL_CONFIG.SHEETS.EMAIL_STATS);
  }
  statsSheet.clear();

  const statsHeaders = ['Date', 'Emails Received', 'Emails Sent', 'Unread Count', 'Lead Emails', 'Response Rate'];
  statsSheet.getRange(1, 1, 1, statsHeaders.length).setValues([statsHeaders]);
  statsSheet.getRange(1, 1, 1, statsHeaders.length)
    .setBackground('#34a853')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  statsSheet.setFrozenRows(1);

  // Templates sheet
  let templatesSheet = ss.getSheetByName(GMAIL_CONFIG.SHEETS.TEMPLATES);
  if (!templatesSheet) {
    templatesSheet = ss.insertSheet(GMAIL_CONFIG.SHEETS.TEMPLATES);
  }
  templatesSheet.clear();

  const templateHeaders = ['Template Name', 'Subject', 'Body', 'Use Case'];
  const defaultTemplates = [
    templateHeaders,
    ['Quick Acknowledgment', 'RE: {{Original Subject}}', 'Hi {{Name}},\n\nThank you for reaching out! I received your message and will get back to you within 24 hours.\n\nBest regards,\n{{COMPANY_NAME}}', 'Auto-acknowledge inquiries'],
    ['Out of Office', 'Out of Office', 'Thank you for your email. I am currently out of the office and will respond upon my return.\n\nFor urgent matters, please call {{COMPANY_PHONE}}.\n\nBest regards,\n{{COMPANY_NAME}}', 'Vacation/away responses'],
    ['Showing Follow-up', 'Great meeting you today!', 'Hi {{Name}},\n\nIt was wonderful showing you properties today! I wanted to follow up and see if you had any questions about what we saw.\n\nLet me know your thoughts when you have a chance.\n\nBest regards,\n{{COMPANY_NAME}}', 'Post-showing follow-up']
  ];

  templatesSheet.getRange(1, 1, defaultTemplates.length, 4).setValues(defaultTemplates);
  templatesSheet.getRange(1, 1, 1, 4)
    .setBackground('#f57c00')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  templatesSheet.setColumnWidth(3, 500);
  templatesSheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert('✅ Gmail automation sheets initialized!');
}

// ============================================
// Daily Digest
// ============================================
function generateDailyDigest() {
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const dateStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // Get emails from last 24 hours
  const threads = GmailApp.search(`after:${dateStr}`, 0, 100);

  let unread = 0;
  let fromClients = [];
  let leads = [];
  let urgent = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const from = lastMessage.getFrom();
    const subject = lastMessage.getSubject();
    const isUnread = thread.isUnread();

    if (isUnread) unread++;

    // Check if from lead source
    const fromLower = from.toLowerCase();
    const subjectLower = subject.toLowerCase();

    for (const pattern of GMAIL_CONFIG.LEAD_PATTERNS.from) {
      if (fromLower.includes(pattern)) {
        leads.push({ from, subject });
        break;
      }
    }

    for (const pattern of GMAIL_CONFIG.LEAD_PATTERNS.subject) {
      if (subjectLower.includes(pattern)) {
        leads.push({ from, subject });
        break;
      }
    }

    // Check for urgent keywords
    if (subjectLower.includes('urgent') || subjectLower.includes('asap') || subjectLower.includes('immediately')) {
      urgent.push({ from, subject });
    }
  });

  // Create digest
  let digest = `📧 DAILY EMAIL DIGEST\n`;
  digest += `${now.toLocaleDateString()}\n\n`;
  digest += `━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  digest += `📊 SUMMARY\n`;
  digest += `• Total emails (24h): ${threads.length}\n`;
  digest += `• Unread: ${unread}\n`;
  digest += `• Potential leads: ${leads.length}\n`;
  digest += `• Urgent: ${urgent.length}\n\n`;

  if (urgent.length > 0) {
    digest += `🚨 URGENT EMAILS\n`;
    urgent.forEach(e => {
      digest += `• ${e.subject}\n  From: ${e.from}\n`;
    });
    digest += `\n`;
  }

  if (leads.length > 0) {
    digest += `🎯 POTENTIAL LEADS\n`;
    leads.slice(0, 10).forEach(e => {
      digest += `• ${e.subject}\n  From: ${e.from}\n`;
    });
    if (leads.length > 10) {
      digest += `  ... and ${leads.length - 10} more\n`;
    }
    digest += `\n`;
  }

  digest += `\n📬 View your inbox: https://mail.google.com`;

  // Show digest or send email
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Daily Digest', digest + '\n\nWould you like to email this digest to yourself?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    const email = Session.getActiveUser().getEmail();
    MailApp.sendEmail({
      to: email,
      subject: `📧 Daily Email Digest - ${now.toLocaleDateString()}`,
      body: digest
    });
    ui.alert('Digest sent to ' + email);
  }
}

function createDigestTrigger() {
  // Remove existing
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyDigestAuto') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create daily trigger at 7 AM
  ScriptApp.newTrigger('sendDailyDigestAuto')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('✅ Daily digest scheduled for 7 AM');
}

function sendDailyDigestAuto() {
  // Same as generateDailyDigest but auto-sends
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const dateStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const threads = GmailApp.search(`after:${dateStr}`, 0, 100);

  let unread = 0;
  let leads = [];
  let urgent = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const lastMessage = messages[messages.length - 1];
    const from = lastMessage.getFrom();
    const subject = lastMessage.getSubject();

    if (thread.isUnread()) unread++;

    const fromLower = from.toLowerCase();
    const subjectLower = subject.toLowerCase();

    for (const pattern of GMAIL_CONFIG.LEAD_PATTERNS.from) {
      if (fromLower.includes(pattern)) {
        leads.push({ from, subject });
        break;
      }
    }

    if (subjectLower.includes('urgent') || subjectLower.includes('asap')) {
      urgent.push({ from, subject });
    }
  });

  let digest = `📧 DAILY EMAIL DIGEST - ${now.toLocaleDateString()}\n\n`;
  digest += `📊 SUMMARY\n`;
  digest += `• Total emails (24h): ${threads.length}\n`;
  digest += `• Unread: ${unread}\n`;
  digest += `• Potential leads: ${leads.length}\n`;
  digest += `• Urgent: ${urgent.length}\n\n`;

  if (urgent.length > 0) {
    digest += `🚨 URGENT EMAILS\n`;
    urgent.forEach(e => digest += `• ${e.subject}\n`);
    digest += `\n`;
  }

  if (leads.length > 0) {
    digest += `🎯 POTENTIAL LEADS\n`;
    leads.slice(0, 5).forEach(e => digest += `• ${e.subject}\n`);
    digest += `\n`;
  }

  const email = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: email,
    subject: `📧 Daily Email Digest - ${now.toLocaleDateString()}`,
    body: digest
  });

  Logger.log('Daily digest sent');
}

// ============================================
// Lead Capture
// ============================================
function captureLeadsFromEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GMAIL_CONFIG.SHEETS.CAPTURED_LEADS);

  if (!sheet) {
    initializeGmailSheets();
    sheet = ss.getSheetByName(GMAIL_CONFIG.SHEETS.CAPTURED_LEADS);
  }

  // Search for potential lead emails from last 7 days
  const searchQueries = [
    ...GMAIL_CONFIG.LEAD_PATTERNS.from.map(p => `from:${p}`),
    ...GMAIL_CONFIG.LEAD_PATTERNS.subject.map(p => `subject:"${p}"`)
  ];

  const query = `newer_than:7d (${searchQueries.join(' OR ')})`;
  const threads = GmailApp.search(query, 0, 50);

  // Get existing captured emails to avoid duplicates
  const existingData = sheet.getDataRange().getValues();
  const existingSubjects = new Set(existingData.slice(1).map(row => row[2]));

  const newLeads = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const firstMessage = messages[0];
    const subject = firstMessage.getSubject();

    // Skip if already captured
    if (existingSubjects.has(subject)) return;

    const from = firstMessage.getFrom();
    const body = firstMessage.getPlainBody();
    const date = firstMessage.getDate();

    // Try to extract info from body
    const extracted = extractLeadInfo(body);

    // Determine source
    let source = 'Unknown';
    for (const pattern of GMAIL_CONFIG.LEAD_PATTERNS.from) {
      if (from.toLowerCase().includes(pattern)) {
        source = pattern.split('.')[0].charAt(0).toUpperCase() + pattern.split('.')[0].slice(1);
        break;
      }
    }

    newLeads.push([
      date,
      from,
      subject,
      source,
      extracted.name || '',
      extracted.email || '',
      extracted.phone || '',
      extracted.property || '',
      body.substring(0, 500),  // Preview
      'New',
      firstMessage.getThread().getPermalink()
    ]);
  });

  // Write new leads
  if (newLeads.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newLeads.length, newLeads[0].length).setValues(newLeads);
  }

  SpreadsheetApp.getUi().alert(`✅ Lead capture complete!\n\nNew leads found: ${newLeads.length}\n\nCheck the "${GMAIL_CONFIG.SHEETS.CAPTURED_LEADS}" sheet.`);
}

function extractLeadInfo(body) {
  const info = {
    name: '',
    email: '',
    phone: '',
    property: ''
  };

  // Email pattern
  const emailMatch = body.match(/[\w.-]+@[\w.-]+\.\w+/);
  if (emailMatch) info.email = emailMatch[0];

  // Phone pattern (various formats)
  const phoneMatch = body.match(/(\+?1?[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/);
  if (phoneMatch) info.phone = phoneMatch[0];

  // Try to extract name (usually after "Name:" or "From:")
  const nameMatch = body.match(/(?:Name|From|Contact):\s*([A-Za-z]+\s+[A-Za-z]+)/i);
  if (nameMatch) info.name = nameMatch[1];

  // Property address pattern
  const addressMatch = body.match(/\d+\s+[\w\s]+(?:Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Lane|Ln|Way|Boulevard|Blvd)/i);
  if (addressMatch) info.property = addressMatch[0];

  return info;
}

function createLeadCaptureTrigger() {
  // Remove existing
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'captureLeadsAuto') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create trigger every 2 hours
  ScriptApp.newTrigger('captureLeadsAuto')
    .timeBased()
    .everyHours(2)
    .create();

  SpreadsheetApp.getUi().alert('✅ Lead capture scheduled to run every 2 hours');
}

function captureLeadsAuto() {
  // Silent version of captureLeadsFromEmail
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GMAIL_CONFIG.SHEETS.CAPTURED_LEADS);

  if (!sheet) return;

  const searchQueries = [
    ...GMAIL_CONFIG.LEAD_PATTERNS.from.map(p => `from:${p}`),
    ...GMAIL_CONFIG.LEAD_PATTERNS.subject.map(p => `subject:"${p}"`)
  ];

  const query = `newer_than:1d (${searchQueries.join(' OR ')})`;
  const threads = GmailApp.search(query, 0, 20);

  const existingData = sheet.getDataRange().getValues();
  const existingSubjects = new Set(existingData.slice(1).map(row => row[2]));

  const newLeads = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const firstMessage = messages[0];
    const subject = firstMessage.getSubject();

    if (existingSubjects.has(subject)) return;

    const from = firstMessage.getFrom();
    const body = firstMessage.getPlainBody();
    const date = firstMessage.getDate();
    const extracted = extractLeadInfo(body);

    let source = 'Unknown';
    for (const pattern of GMAIL_CONFIG.LEAD_PATTERNS.from) {
      if (from.toLowerCase().includes(pattern)) {
        source = pattern.split('.')[0].charAt(0).toUpperCase() + pattern.split('.')[0].slice(1);
        break;
      }
    }

    newLeads.push([
      date, from, subject, source,
      extracted.name || '', extracted.email || '', extracted.phone || '', extracted.property || '',
      body.substring(0, 500), 'New', firstMessage.getThread().getPermalink()
    ]);
  });

  if (newLeads.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newLeads.length, newLeads[0].length).setValues(newLeads);
    Logger.log('Captured ' + newLeads.length + ' new leads');
  }
}

// ============================================
// Statistics
// ============================================
function showEmailStatistics() {
  const now = new Date();
  const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
  const dateStr = Utilities.formatDate(thirtyDaysAgo, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // Count received
  const received = GmailApp.search(`after:${dateStr} in:anywhere`);

  // Count sent
  const sent = GmailApp.search(`after:${dateStr} in:sent`);

  // Count unread
  const unread = GmailApp.search('is:unread');

  // Count by category
  const primary = GmailApp.search(`after:${dateStr} category:primary`);
  const social = GmailApp.search(`after:${dateStr} category:social`);
  const promotions = GmailApp.search(`after:${dateStr} category:promotions`);

  let report = `📊 EMAIL STATISTICS (Last 30 Days)\n\n`;
  report += `━━━━━━━━━━━━━━━━━━━━━━━━\n\n`;
  report += `📬 VOLUME\n`;
  report += `• Received: ${received.length}\n`;
  report += `• Sent: ${sent.length}\n`;
  report += `• Currently unread: ${unread.length}\n\n`;
  report += `📁 BY CATEGORY\n`;
  report += `• Primary: ${primary.length}\n`;
  report += `• Social: ${social.length}\n`;
  report += `• Promotions: ${promotions.length}\n\n`;

  const avgPerDay = Math.round(received.length / 30);
  report += `📈 AVERAGES\n`;
  report += `• Emails per day: ~${avgPerDay}\n`;
  report += `• Sent per day: ~${Math.round(sent.length / 30)}\n`;

  SpreadsheetApp.getUi().alert(report);
}

function showUnreadSummary() {
  const threads = GmailApp.search('is:unread', 0, 50);

  let report = `📬 UNREAD EMAILS (${threads.length} total)\n\n`;

  if (threads.length === 0) {
    report += 'Inbox Zero! Great job! 🎉';
  } else {
    // Group by age
    const today = new Date();
    let todayCount = 0;
    let weekCount = 0;
    let olderCount = 0;

    threads.forEach(thread => {
      const date = thread.getLastMessageDate();
      const daysDiff = Math.floor((today - date) / (1000 * 60 * 60 * 24));

      if (daysDiff === 0) todayCount++;
      else if (daysDiff <= 7) weekCount++;
      else olderCount++;
    });

    report += `📅 BY AGE\n`;
    report += `• Today: ${todayCount}\n`;
    report += `• This week: ${weekCount}\n`;
    report += `• Older: ${olderCount}\n\n`;

    report += `📋 RECENT UNREAD\n`;
    threads.slice(0, 10).forEach(thread => {
      const subject = thread.getFirstMessageSubject();
      report += `• ${subject.substring(0, 50)}${subject.length > 50 ? '...' : ''}\n`;
    });
  }

  SpreadsheetApp.getUi().alert(report);
}

function showResponseTimeReport() {
  SpreadsheetApp.getUi().alert('Response time analysis requires more processing.\n\nThis feature tracks how quickly you respond to emails - coming soon!');
}

// ============================================
// Cleanup Utilities
// ============================================
function archiveOldNewsletters() {
  const days = GMAIL_CONFIG.AUTO_ARCHIVE.newsletters;
  const query = `category:updates older_than:${days}d is:unread`;

  const threads = GmailApp.search(query, 0, 100);
  let count = 0;

  threads.forEach(thread => {
    thread.markRead();
    thread.moveToArchive();
    count++;
  });

  SpreadsheetApp.getUi().alert(`✅ Archived ${count} newsletter emails older than ${days} days.`);
}

function archiveOldPromotions() {
  const days = GMAIL_CONFIG.AUTO_ARCHIVE.promotions;
  const query = `category:promotions older_than:${days}d`;

  const threads = GmailApp.search(query, 0, 100);
  let count = 0;

  threads.forEach(thread => {
    thread.markRead();
    thread.moveToArchive();
    count++;
  });

  SpreadsheetApp.getUi().alert(`✅ Archived ${count} promotional emails older than ${days} days.`);
}

function emptyTrash() {
  // Note: Gmail API doesn't directly support emptying trash
  // This marks old trash items as permanently deleted
  SpreadsheetApp.getUi().alert('To empty trash, go to Gmail > Trash > "Empty Trash now"\n\nDirect deletion via script requires Gmail API.');
}
