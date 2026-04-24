#!/usr/bin/env node
/**
 * Workspace Tools - Client Setup Script
 * Creates a standalone Google Apps Script project for a client
 *
 * Usage:
 *   node setup.js                     # Interactive mode
 *   node setup.js --config FILE       # Config file mode (for batch deployment)
 *
 * True North Data Strategies
 * jacob@truenorthstrategyops.com
 */

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { execSync } = require('child_process');

// Check for config file mode (batch deployment)
const args = process.argv.slice(2);
const configIndex = args.indexOf('--config');
let batchConfig = null;

if (configIndex !== -1 && args[configIndex + 1]) {
  const configPath = path.resolve(args[configIndex + 1]);
  if (fs.existsSync(configPath)) {
    batchConfig = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    console.log('  📄 Using config file for batch deployment');
  }
}

// Industry Templates - folder structures AND Gmail labels for different business types
const INDUSTRY_TEMPLATES = {
  'real-estate': {
    name: 'Real Estate',
    emoji: '🏠',
    folders: {
      'Listings': ['Active', 'Pending', 'Sold', 'Coming Soon', 'Photos & Media'],
      'Clients': ['Buyers', 'Sellers', 'Prospects', 'Past Clients'],
      'Transactions': ['2025', '2026', 'Templates', 'Disclosures'],
      'Marketing': ['Flyers', 'Social Media', 'Branding', 'Email Templates'],
      'Operations': ['Policies', 'Training', 'Vendor Contacts', 'Licenses & Docs']
    },
    labels: {
      'Clients': ['Buyers', 'Sellers', 'Prospects', 'Past Clients'],
      'Transactions': ['Active', 'Pending', 'Closed'],
      'Listings': ['Active', 'Coming Soon', 'Sold'],
      'Leads': ['Website', 'Zillow', 'Realtor.com', 'Referral', 'Open House'],
      'Vendors': ['Title', 'Lenders', 'Inspectors', 'Contractors'],
      'Marketing': [],
      'Admin': ['Brokerage', 'MLS', 'License'],
      'Action Required': [],
      'Waiting For Response': []
    },
    autoLabelRules: [
      { label: 'Leads/Zillow', from: ['zillow.com', 'zillowgroup.com'], subject: [] },
      { label: 'Leads/Realtor.com', from: ['realtor.com', 'move.com'], subject: [] },
      { label: 'Leads/Website', from: [], subject: ['website inquiry', 'contact form', 'new lead'] },
      { label: 'Vendors/Title', from: [], subject: ['title', 'closing', 'escrow'] },
      { label: 'Vendors/Lenders', from: [], subject: ['mortgage', 'loan', 'pre-approval'] }
    ]
  },
  'consulting': {
    name: 'Consulting / Agency',
    emoji: '💼',
    folders: {
      'Clients': ['Active', 'Prospects', 'Archive'],
      'Projects': ['Active', 'Completed', 'Templates'],
      'Proposals': ['Templates', 'Sent', 'Archive'],
      'Finance': ['Invoices', 'Expenses', 'Reports'],
      'Operations': ['Contracts', 'SOPs', 'Training']
    },
    labels: {
      'Clients': ['Active', 'Prospects', 'Past'],
      'Projects': ['Active', 'Pending', 'Completed'],
      'Proposals': ['Draft', 'Sent', 'Won', 'Lost'],
      'Invoicing': ['Pending', 'Paid', 'Overdue'],
      'Vendors': ['Contractors', 'Tools', 'Services'],
      'Marketing': [],
      'Admin': ['Contracts', 'Legal', 'Insurance'],
      'Action Required': [],
      'Waiting For Response': []
    },
    autoLabelRules: [
      { label: 'Invoicing/Pending', from: [], subject: ['invoice', 'payment due'] },
      { label: 'Proposals/Sent', from: [], subject: ['proposal', 'quote', 'estimate'] },
      { label: 'Projects/Active', from: [], subject: ['project update', 'deliverable'] }
    ]
  },
  'construction': {
    name: 'Construction / Trades',
    emoji: '🔧',
    folders: {
      'Jobs': ['Active', 'Completed', 'Bidding', 'Photos'],
      'Clients': ['Residential', 'Commercial', 'Archive'],
      'Estimates': ['Templates', 'Sent', 'Won'],
      'Safety': ['Training', 'Certifications', 'Incidents'],
      'Operations': ['Contracts', 'Insurance', 'Licenses']
    },
    labels: {
      'Jobs': ['Active', 'Bidding', 'Completed'],
      'Clients': ['Residential', 'Commercial', 'Repeat'],
      'Subcontractors': ['Electrical', 'Plumbing', 'HVAC', 'General'],
      'Suppliers': ['Materials', 'Equipment', 'Rentals'],
      'Permits': ['Pending', 'Approved', 'Inspections'],
      'Admin': ['Insurance', 'Licenses', 'Safety'],
      'Action Required': [],
      'Waiting For Response': []
    },
    autoLabelRules: [
      { label: 'Permits/Pending', from: [], subject: ['permit', 'building dept', 'inspection'] },
      { label: 'Jobs/Bidding', from: [], subject: ['bid', 'rfp', 'quote request'] },
      { label: 'Suppliers/Materials', from: [], subject: ['order', 'delivery', 'shipment'] }
    ]
  },
  'healthcare': {
    name: 'Healthcare / Medical',
    emoji: '🏥',
    folders: {
      'Patients': ['Active', 'Archive', 'Forms'],
      'Insurance': ['Claims', 'EOBs', 'Authorizations'],
      'Compliance': ['HIPAA', 'Training', 'Audits'],
      'Operations': ['Policies', 'Procedures', 'Staff'],
      'Marketing': ['Website', 'Social', 'Reviews']
    },
    labels: {
      'Patients': ['Active', 'New', 'Archive'],
      'Appointments': ['Scheduled', 'Confirmed', 'Cancelled'],
      'Insurance': ['Claims', 'Authorizations', 'Denials'],
      'Referrals': ['Incoming', 'Outgoing'],
      'Vendors': ['Labs', 'Imaging', 'Specialists', 'Supplies'],
      'Admin': ['Compliance', 'Licensing', 'Training'],
      'Action Required': [],
      'Urgent': []
    },
    autoLabelRules: [
      { label: 'Insurance/Claims', from: [], subject: ['claim', 'eob', 'remittance'] },
      { label: 'Referrals/Incoming', from: [], subject: ['referral', 'patient transfer'] },
      { label: 'Appointments/Scheduled', from: [], subject: ['appointment', 'scheduling'] }
    ]
  },
  'ecommerce': {
    name: 'E-Commerce / Retail',
    emoji: '🛒',
    folders: {
      'Products': ['Active', 'Discontinued', 'Photos'],
      'Orders': ['Processing', 'Archive', 'Returns'],
      'Marketing': ['Campaigns', 'Assets', 'Analytics'],
      'Finance': ['Invoices', 'Tax', 'Reports'],
      'Operations': ['Suppliers', 'Shipping', 'Inventory']
    },
    labels: {
      'Orders': ['New', 'Processing', 'Shipped', 'Returns'],
      'Customers': ['VIP', 'Repeat', 'Support'],
      'Suppliers': ['Inventory', 'Dropship', 'Returns'],
      'Marketing': ['Campaigns', 'Social', 'Affiliates'],
      'Finance': ['Invoices', 'Refunds', 'Disputes'],
      'Admin': ['Legal', 'Compliance', 'Tax'],
      'Action Required': [],
      'Urgent': []
    },
    autoLabelRules: [
      { label: 'Orders/New', from: [], subject: ['new order', 'order confirmation'] },
      { label: 'Orders/Returns', from: [], subject: ['return request', 'refund', 'rma'] },
      { label: 'Customers/Support', from: [], subject: ['support ticket', 'help request'] }
    ]
  },
  'general': {
    name: 'General Business',
    emoji: '🏢',
    folders: {
      'Clients': ['Active', 'Prospects', 'Archive'],
      'Projects': ['Active', 'Completed', 'Templates'],
      'Finance': ['Invoices', 'Expenses', 'Reports'],
      'Marketing': ['Assets', 'Campaigns'],
      'Operations': ['Contracts', 'Policies', 'HR']
    },
    labels: {
      'Clients': ['Active', 'Prospects', 'Past'],
      'Projects': ['Active', 'Completed'],
      'Finance': ['Invoices', 'Expenses', 'Reports'],
      'Vendors': ['Suppliers', 'Services'],
      'Marketing': [],
      'Admin': ['Contracts', 'Legal', 'HR'],
      'Action Required': [],
      'Waiting For Response': []
    },
    autoLabelRules: [
      { label: 'Finance/Invoices', from: [], subject: ['invoice', 'payment'] },
      { label: 'Projects/Active', from: [], subject: ['project', 'update'] }
    ]
  },
  'analytics': {
    name: 'Analytics / Data',
    emoji: '📊',
    folders: {
      'Clients': ['Active', 'Prospects', 'Archive'],
      'Projects': ['Active', 'Completed', 'Templates'],
      'Data': ['Raw', 'Processed', 'Reports'],
      'Dashboards': ['Templates', 'Client Dashboards'],
      'Operations': ['Contracts', 'SOPs', 'Training']
    },
    labels: {
      'Clients': ['Active', 'Prospects', 'Past'],
      'Projects': ['Active', 'Completed', 'On Hold'],
      'Data': ['Requests', 'Deliveries'],
      'Dashboards': ['Active', 'Archive'],
      'Vendors': ['Data Sources', 'Tools'],
      'Admin': ['Contracts', 'SOPs'],
      'Action Required': [],
      'Waiting For Response': []
    },
    autoLabelRules: [
      { label: 'Data/Requests', from: [], subject: ['data request', 'report request'] },
      { label: 'Projects/Active', from: [], subject: ['dashboard', 'analytics'] }
    ]
  },
  'fleet': {
    name: 'Fleet / Automotive',
    emoji: '🚗',
    folders: {
      'Vehicles': ['Active', 'Maintenance', 'Archive'],
      'Drivers': ['Active', 'Training', 'Records'],
      'Maintenance': ['Scheduled', 'Completed', 'Vendors'],
      'Compliance': ['Inspections', 'Registrations', 'Insurance'],
      'Operations': ['Routes', 'Fuel', 'Reports']
    },
    labels: {
      'Vehicles': ['Active', 'Maintenance', 'Retired'],
      'Drivers': ['Active', 'Training', 'Incidents'],
      'Maintenance': ['Scheduled', 'Urgent', 'Completed'],
      'Compliance': ['Inspections', 'Registrations', 'Insurance'],
      'Vendors': ['Mechanics', 'Parts', 'Fuel'],
      'Admin': ['DOT', 'Legal'],
      'Action Required': [],
      'Urgent': []
    },
    autoLabelRules: [
      { label: 'Maintenance/Scheduled', from: [], subject: ['maintenance', 'service due'] },
      { label: 'Compliance/Inspections', from: [], subject: ['inspection', 'dot', 'compliance'] }
    ]
  },
  'strategy': {
    name: 'Strategy / Premium',
    emoji: '⭐',
    folders: {
      'Clients': ['Active', 'Prospects', 'Archive'],
      'Engagements': ['Active', 'Completed', 'Templates'],
      'Deliverables': ['Reports', 'Presentations', 'Frameworks'],
      'Research': ['Market', 'Competitive', 'Industry'],
      'Operations': ['Contracts', 'Proposals', 'Admin']
    },
    labels: {
      'Clients': ['Active', 'Prospects', 'Past'],
      'Engagements': ['Active', 'Pipeline', 'Completed'],
      'Proposals': ['Draft', 'Sent', 'Won', 'Lost'],
      'Research': ['Market', 'Competitive'],
      'Vendors': ['Partners', 'Contractors'],
      'Admin': ['Contracts', 'Legal', 'Finance'],
      'Action Required': [],
      'Waiting For Response': []
    },
    autoLabelRules: [
      { label: 'Proposals/Sent', from: [], subject: ['proposal', 'engagement letter'] },
      { label: 'Engagements/Active', from: [], subject: ['project update', 'deliverable'] },
      { label: 'Research/Competitive', from: [], subject: ['competitor', 'market intel'] }
    ]
  },
  'small-business': {
    name: 'Small Business',
    emoji: '🏪',
    folders: {
      'Customers': ['Active', 'Prospects', 'Past'],
      'Sales': ['Quotes', 'Orders', 'Invoices'],
      'Products-Services': ['Active', 'Discontinued', 'Pricing'],
      'Marketing': ['Social Media', 'Ads', 'Email Campaigns'],
      'Finance': ['Income', 'Expenses', 'Tax Documents', 'Bank Statements'],
      'Operations': ['Policies', 'Procedures', 'Vendor Contracts', 'Licenses']
    },
    labels: {
      'Customers': ['Active', 'Prospects', 'Past', 'VIP'],
      'Sales': ['New Orders', 'Pending', 'Completed', 'Returns'],
      'Vendors': ['Suppliers', 'Services', 'Contractors'],
      'Finance': ['Invoices', 'Payments', 'Expenses', 'Tax'],
      'Marketing': ['Campaigns', 'Inquiries'],
      'Admin': ['Insurance', 'Legal', 'Licenses'],
      'Action Required': [],
      'Follow Up': []
    },
    autoLabelRules: [
      { label: 'Sales/New Orders', from: [], subject: ['new order', 'purchase', 'order confirmation'] },
      { label: 'Finance/Invoices', from: [], subject: ['invoice', 'bill', 'payment due'] },
      { label: 'Customers/Prospects', from: [], subject: ['inquiry', 'question about', 'interested in'] },
      { label: 'Finance/Payments', from: [], subject: ['payment received', 'payment confirmation'] }
    ]
  },
  'developer': {
    name: 'Developer / Tech',
    emoji: '👨‍💻',
    folders: {
      'Projects': ['Active', 'Archived', 'Ideas', 'Templates'],
      'Clients': ['Active', 'Prospects', 'Completed'],
      'Documentation': ['API Docs', 'Technical Specs', 'Architecture'],
      'Learning': ['Courses', 'Tutorials', 'Certifications', 'Notes'],
      'Assets': ['Icons', 'Images', 'Fonts', 'Design Files'],
      'DevOps': ['Configs', 'Scripts', 'Credentials', 'Backups']
    },
    labels: {
      'Projects': ['Active', 'Backlog', 'Archived'],
      'Clients': ['Active', 'Prospects', 'Completed'],
      'GitHub': ['PRs', 'Issues', 'Releases', 'Actions'],
      'CI-CD': ['Builds', 'Deployments', 'Failures'],
      'Security': ['Alerts', 'Vulnerabilities', 'Updates'],
      'Newsletters': ['Dev News', 'Tools', 'Frameworks'],
      'Billing': ['Hosting', 'Services', 'Subscriptions'],
      'Action Required': [],
      'Review Needed': []
    },
    autoLabelRules: [
      { label: 'GitHub/PRs', from: ['github.com', 'notifications@github.com'], subject: ['pull request', 'merged', 'review requested'] },
      { label: 'GitHub/Issues', from: ['github.com'], subject: ['issue', 'bug', 'feature request'] },
      { label: 'GitHub/Actions', from: ['github.com'], subject: ['workflow', 'action', 'run failed', 'run succeeded'] },
      { label: 'CI-CD/Failures', from: ['vercel.com', 'netlify.com', 'railway.app', 'render.com'], subject: ['failed', 'error', 'deployment failed'] },
      { label: 'Security/Alerts', from: ['dependabot', 'snyk.io', 'github.com'], subject: ['security', 'vulnerability', 'cve', 'security alert'] },
      { label: 'Billing/Hosting', from: ['aws.amazon.com', 'cloud.google.com', 'azure.com', 'digitalocean.com', 'vercel.com'], subject: ['invoice', 'billing', 'payment'] }
    ]
  }
};

// Emoji to industry default mapping
const EMOJI_INDUSTRY_MAP = {
  '🏠': 'real-estate',
  '🏢': 'general',
  '🔧': 'construction',
  '⚙️': 'developer',
  '📊': 'analytics',
  '🚗': 'fleet',
  '💼': 'consulting',
  '🧭': 'strategy',
  '⭐': 'strategy',
  '🛠️': 'construction',
  '🏥': 'healthcare',
  '🛒': 'ecommerce',
  '🏪': 'small-business',
  '👨‍💻': 'developer',
  '💻': 'developer',
  '🖥️': 'developer'
};

// Module Configuration
const MODULE_CONFIG = {
  name: 'Operations Toolkit',
  description: 'Workspace automation tools (Drive, Gmail, Calendar)',
  projectSuffix: 'Workspace-Tools',
  projectType: 'sheets',
  defaultIcon: '🏢',
  defaultIndustry: 'general',

  // Files to include in deployment
  templateFiles: [
    'core/config-module.gs',
    'core/WorkspaceTools.gs',
    'core/drive-folder-setup.gs',
    'core/drive-inventory.gs',
    'core/drive-module.gs',
    'core/gmail-label-creator.gs',
    'core/gmail-automation.gs',
    'core/gmail-module.gs',
    'core/logger.gs',
    'core/runner.gs',
    'ui/Sidebar.html',
    'ui/Dashboard.html',
    'ui/Help.html',
    'ui/UserManual.html',
    'ui/presentation.html',
    'appsscript.json',
    '.claspignore'
  ],

  // File rename mappings (source -> destination pattern)
  fileRenames: {},

  // Legacy string replacements (none needed - source files use proper placeholders)
  legacyReplacements: {},

  // Next steps displayed after setup
  nextSteps: [
    'Open the Google Sheet that was created',
    'Refresh the page to see the custom menu',
    'Click the custom menu in the toolbar',
    'Go to Setup & Initialize > Initialize All Sheets',
    'Grant permissions when prompted'
  ]
};

// =========================================
// Readline Utilities
// =========================================

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

function ask(question) {
  return new Promise(resolve => rl.question(question, resolve));
}

// =========================================
// Main Setup Function
// =========================================

async function main() {
  console.log('\n========================================');
  console.log(`   ${MODULE_CONFIG.name} - Client Setup`);
  console.log(`   ${MODULE_CONFIG.description}`);
  console.log('========================================\n');
  console.log('This wizard will create a customized Google Apps Script project');
  console.log('with your company branding and deploy it to Google Workspace.\n');

  // Gather client information
  const answers = await collectCompanyInfo();

  // Display summary
  displaySummary(answers);

  // Confirm (skip in batch mode)
  if (!batchConfig) {
    const confirm = await ask('\nProceed with setup? (y/n): ');
    if (confirm.toLowerCase() !== 'y') {
      console.log('Setup cancelled.');
      rl.close();
      process.exit(0);
    }
  }

  // Create output directory
  const safeName = createSafeName(answers.companyName);
  const outputDir = await createOutputDirectory(safeName);

  // Process template files
  console.log('\nProcessing template files...');
  const replacements = buildReplacementMap(answers, safeName);
  processFiles(outputDir, replacements);

  // Create .clasp.json.template for reference
  createClaspTemplate(outputDir);

  // Close readline before clasp operations
  rl.close();

  // Deploy with clasp
  const projectTitle = `${answers.shortName} ${MODULE_CONFIG.projectSuffix}`;
  const claspJson = deployWithClasp(outputDir, projectTitle);

  // Display completion message
  displayComplete(claspJson, answers);
}

// =========================================
// Company Info Collection
// =========================================

async function collectCompanyInfo() {
  const answers = {};

  // If batch config is provided, use it instead of prompting
  if (batchConfig) {
    answers.companyName = batchConfig.companyName;
    answers.shortName = batchConfig.shortName || batchConfig.companyName;
    answers.menuIcon = batchConfig.menuIcon || MODULE_CONFIG.defaultIcon;
    answers.industry = batchConfig.industryKey || MODULE_CONFIG.defaultIndustry;
    answers.industryData = INDUSTRY_TEMPLATES[answers.industry] || INDUSTRY_TEMPLATES[MODULE_CONFIG.defaultIndustry];
    answers.companyEmail = batchConfig.companyEmail || '';
    answers.companyPhone = batchConfig.companyPhone || '';
    answers.companyWebsite = batchConfig.companyWebsite || '';

    console.log(`\n  Company: ${answers.companyName}`);
    console.log(`  Industry: ${answers.industryData.emoji} ${answers.industryData.name}`);
    return answers;
  }

  // Interactive mode - prompt for all values
  // Required: Company Name
  answers.companyName = await ask('Company Name (e.g., "Acme Corp"): ');
  if (!answers.companyName.trim()) {
    console.log('Error: Company name is required.');
    rl.close();
    process.exit(1);
  }
  answers.companyName = answers.companyName.trim();

  // Optional: Short Name (default to company name)
  const shortNameInput = await ask(`Short Name for menus [${answers.companyName}]: `);
  answers.shortName = shortNameInput.trim() || answers.companyName;

  // Optional: Menu Icon
  const iconInput = await ask(`Menu Icon emoji [${MODULE_CONFIG.defaultIcon}]: `);
  answers.menuIcon = iconInput.trim() || MODULE_CONFIG.defaultIcon;

  // Industry Selection - show options based on emoji or let user choose
  const suggestedIndustry = EMOJI_INDUSTRY_MAP[answers.menuIcon] || MODULE_CONFIG.defaultIndustry;
  console.log('\n--- Industry Folder Templates ---');
  const industries = Object.keys(INDUSTRY_TEMPLATES);
  industries.forEach((key, i) => {
    const ind = INDUSTRY_TEMPLATES[key];
    const marker = key === suggestedIndustry ? ' (suggested)' : '';
    console.log(`  ${i + 1}. ${ind.emoji} ${ind.name}${marker}`);
  });
  const industryInput = await ask(`\nSelect industry (1-${industries.length}) [${industries.indexOf(suggestedIndustry) + 1}]: `);
  const industryIndex = parseInt(industryInput.trim()) - 1;
  if (industryIndex >= 0 && industryIndex < industries.length) {
    answers.industry = industries[industryIndex];
  } else {
    answers.industry = suggestedIndustry;
  }
  answers.industryData = INDUSTRY_TEMPLATES[answers.industry];

  // Optional: Email
  answers.companyEmail = (await ask('Company Email (optional): ')).trim();

  // Optional: Phone
  answers.companyPhone = (await ask('Company Phone (optional): ')).trim();

  // Optional: Website
  answers.companyWebsite = (await ask('Company Website (optional): ')).trim();

  return answers;
}

// =========================================
// Display Functions
// =========================================

function displaySummary(answers) {
  console.log('\n========================================');
  console.log('         CONFIGURATION SUMMARY');
  console.log('========================================\n');

  console.log('COMPANY INFORMATION:');
  console.log(`  Company Name: ${answers.companyName}`);
  console.log(`  Short Name:   ${answers.shortName}`);
  console.log(`  Menu Icon:    ${answers.menuIcon}`);
  console.log(`  Email:        ${answers.companyEmail || '(not set)'}`);
  console.log(`  Phone:        ${answers.companyPhone || '(not set)'}`);
  console.log(`  Website:      ${answers.companyWebsite || '(not set)'}`);

  console.log('\nINDUSTRY TEMPLATE:');
  console.log(`  ${answers.industryData.emoji} ${answers.industryData.name}`);

  console.log('\nDRIVE FOLDERS TO CREATE:');
  for (const [folder, subfolders] of Object.entries(answers.industryData.folders)) {
    console.log(`  📁 ${folder}/`);
    subfolders.forEach(sub => console.log(`      └── ${sub}`));
  }

  console.log('\nGMAIL LABELS TO CREATE:');
  for (const [label, sublabels] of Object.entries(answers.industryData.labels)) {
    if (sublabels.length > 0) {
      console.log(`  🏷️  ${label}: ${sublabels.join(', ')}`);
    } else {
      console.log(`  🏷️  ${label}`);
    }
  }

  if (answers.industryData.autoLabelRules && answers.industryData.autoLabelRules.length > 0) {
    console.log('\nAUTO-LABEL RULES:');
    answers.industryData.autoLabelRules.forEach(rule => {
      const trigger = rule.from.length > 0 ?
        `from: ${rule.from.join(', ')}` :
        `subject: ${rule.subject.join(', ')}`;
      console.log(`  📨 ${rule.label} ← ${trigger}`);
    });
  }

  console.log('\n========================================');
}

function displayComplete(claspJson, answers) {
  console.log('\n========================================');
  console.log('           SETUP COMPLETE!');
  console.log('========================================\n');
  console.log(`${MODULE_CONFIG.name} created successfully!\n`);

  if (claspJson && claspJson.scriptId) {
    console.log('Script Editor:');
    console.log(`  https://script.google.com/d/${claspJson.scriptId}/edit\n`);
  }

  console.log('Next Steps:');
  MODULE_CONFIG.nextSteps.forEach((step, i) => {
    let displayStep = step;
    if (step.includes('custom menu')) {
      displayStep = `Click "${answers.menuIcon} ${answers.shortName} Tools" menu`;
    }
    console.log(`  ${i + 1}. ${displayStep}`);
  });
  console.log('');

  console.log('Support: jacob@truenorthstrategyops.com\n');
}

// =========================================
// File Processing
// =========================================

function createSafeName(companyName) {
  return companyName
    .replace(/[^a-zA-Z0-9]/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

async function createOutputDirectory(safeName) {
  const dirName = `${safeName}-${MODULE_CONFIG.projectSuffix}`;
  const outputDir = path.join(__dirname, '..', '..', dirName);

  if (fs.existsSync(outputDir)) {
    // Auto-overwrite in batch mode
    if (batchConfig) {
      fs.rmSync(outputDir, { recursive: true });
    } else {
      const overwrite = await ask(`\nDirectory "${dirName}" exists. Overwrite? (y/n): `);
      if (overwrite.toLowerCase() !== 'y') {
        console.log('Setup cancelled.');
        rl.close();
        process.exit(0);
      }
      fs.rmSync(outputDir, { recursive: true });
    }
  }

  console.log(`\nCreating project in: ${outputDir}`);
  fs.mkdirSync(outputDir, { recursive: true });

  return outputDir;
}

// Helper: Build Gmail label config JSON from simple label structure
function buildLabelConfigJSON(labels) {
  const colors = ['#16a765', '#42d692', '#ffad46', '#b3b3b3', '#4285f4', '#fb4c2f', '#a479e2'];
  let colorIndex = 0;

  const labelConfig = {};
  for (const [parent, children] of Object.entries(labels)) {
    const childObj = {};
    if (children && children.length > 0) {
      children.forEach(child => {
        childObj[`${parent}/${child}`] = { color: colors[colorIndex % colors.length] };
        colorIndex++;
      });
    }
    labelConfig[parent] = {
      color: parent === 'Action Required' ? '#fb4c2f' : (parent === 'Waiting For Response' || parent === 'Urgent' ? '#ffad46' : null),
      children: childObj
    };
  }

  return JSON.stringify(labelConfig, null, 4).replace(/\n/g, '\n    ');
}

function buildReplacementMap(answers, safeName) {
  // Escape single quotes to prevent JavaScript string syntax errors
  const escapeForJS = (str) => str ? str.replace(/'/g, "\\'") : '';

  // Build folder structure JSON for drive-folder-setup.gs
  const folderStructure = JSON.stringify(answers.industryData.folders, null, 4)
    .replace(/\n/g, '\n    ');  // Indent for code formatting

  // Build Gmail label structure for gmail-label-creator.gs
  const labelStructure = buildLabelConfigJSON(answers.industryData.labels);
  const autoLabelRules = JSON.stringify(answers.industryData.autoLabelRules || [], null, 4)
    .replace(/\n/g, '\n  ');

  const replacements = {
    // Standard placeholders (escaped for JavaScript strings)
    '{{COMPANY_NAME}}': escapeForJS(answers.companyName),
    '{{SHORT_NAME}}': escapeForJS(answers.shortName),
    '{{MENU_ICON}}': answers.menuIcon,
    '{{COMPANY_EMAIL}}': answers.companyEmail || '',
    '{{COMPANY_PHONE}}': answers.companyPhone || '',
    '{{COMPANY_WEBSITE}}': answers.companyWebsite || '',
    '{{SAFE_NAME}}': safeName,
    '{{TRUENORTH_EMAIL}}': 'jacob@truenorthstrategyops.com',
    '{{INDUSTRY_KEY}}': answers.industry,
    '{{INDUSTRY_NAME}}': answers.industryData.name,
    '{{FOLDER_STRUCTURE}}': folderStructure,
    '{{LABEL_STRUCTURE}}': labelStructure,
    '{{AUTO_LABEL_RULES}}': autoLabelRules
  };

  // Add legacy replacements for backward compatibility (also escaped)
  for (const [search, answerKey] of Object.entries(MODULE_CONFIG.legacyReplacements)) {
    if (answerKey === 'companyName') {
      replacements[search] = escapeForJS(answers.companyName);
    } else if (answerKey === 'shortName') {
      replacements[search] = escapeForJS(answers.shortName);
    } else if (answerKey === 'shortNameTools') {
      replacements[search] = `${escapeForJS(answers.shortName)} Tools`;
    } else if (answerKey === 'menuIcon') {
      replacements[search] = answers.menuIcon;
    } else if (answers[answerKey]) {
      replacements[search] = escapeForJS(answers[answerKey]);
    }
  }

  return replacements;
}

function processFiles(outputDir, replacements) {
  for (const filename of MODULE_CONFIG.templateFiles) {
    const srcPath = path.join(__dirname, '..', filename);
    let destFilename = MODULE_CONFIG.fileRenames[filename] || filename;

    // Apply replacements to filename
    for (const [search, replace] of Object.entries(replacements)) {
      destFilename = destFilename.split(search).join(replace);
    }

    const destPath = path.join(outputDir, destFilename);

    if (fs.existsSync(srcPath)) {
      let content = fs.readFileSync(srcPath, 'utf8');

      // Apply replacements to content
      for (const [search, replace] of Object.entries(replacements)) {
        content = content.split(search).join(replace);
      }

      fs.mkdirSync(path.dirname(destPath), { recursive: true });
      fs.writeFileSync(destPath, content);
      console.log(`  ✓ ${destFilename}`);
    } else {
      console.log(`  ⚠ ${filename} not found, skipping`);
    }
  }
}

function createClaspTemplate(outputDir) {
  const template = {
    scriptId: 'WILL_BE_SET_BY_CLASP_CREATE',
    rootDir: './'
  };
  const templatePath = path.join(outputDir, '.clasp.json.template');
  fs.writeFileSync(templatePath, JSON.stringify(template, null, 2));
  console.log('  ✓ .clasp.json.template');
}

// =========================================
// Clasp Deployment
// =========================================

function deployWithClasp(outputDir, projectTitle) {
  console.log('\n--- Creating Google Apps Script Project ---');

  try {
    process.chdir(outputDir);

    // Check if clasp is available
    try {
      execSync('clasp --version', { stdio: 'pipe' });
    } catch (e) {
      console.log('\nNote: clasp CLI not found or not logged in.');
      console.log('Install with: npm install -g @google/clasp');
      console.log('Login with: clasp login');
      console.log('\nFiles have been prepared in the output directory.');
      console.log('You can manually run these commands later:');
      console.log(`  cd "${outputDir}"`);
      console.log(`  clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`);
      console.log('  clasp push --force');
      return null;
    }

    // Create new project
    console.log('Creating new spreadsheet and script...');
    execSync(`clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`, { stdio: 'inherit' });

    // Push files
    console.log('\nPushing files to Google Apps Script...');
    execSync('clasp push --force', { stdio: 'inherit' });

    // Read the created .clasp.json
    const claspJsonPath = path.join(outputDir, '.clasp.json');
    if (fs.existsSync(claspJsonPath)) {
      return JSON.parse(fs.readFileSync(claspJsonPath, 'utf8'));
    }

    return null;

  } catch (error) {
    console.error('\nError during clasp operations:', error.message);
    console.log('\nFiles have been prepared. You can manually run:');
    console.log(`  cd "${outputDir}"`);
    console.log(`  clasp create --title "${projectTitle}" --type ${MODULE_CONFIG.projectType}`);
    console.log('  clasp push --force');
    return null;
  }
}

// =========================================
// Run
// =========================================

main().catch(err => {
  console.error('Setup failed:', err);
  rl.close();
  process.exit(1);
});
