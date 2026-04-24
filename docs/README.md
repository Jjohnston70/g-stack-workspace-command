# Workspace Tools (Operations Toolkit)

Google Workspace Setup & Organization Suite - Drive structure, Gmail labels, and productivity automation.

## Features

### 📁 Drive Tools
- **Folder Structure**: Create organized folder hierarchy for any business
- **Drive Inventory**: Scan and catalog all files with metadata
- **Duplicate Finder**: Identify potential duplicate files
- **Large/Old Files**: Find files taking up space or needing archive
- **Move Recommendations**: AI-suggested file reorganization

### 📧 Gmail Tools
- **Label Management**: Create hierarchical label structure
- **Auto-Labeling**: Automatically label incoming emails by rules
- **Daily Digest**: Morning summary of email activity
- **Lead Capture**: Extract lead info from portal emails (Zillow, Realtor.com, etc.)
- **Archive Utilities**: Clean up newsletters and promotions

## Quick Start

### Option 1: Automated Setup (Recommended)

From the repo root:

```bash
node setup/setup.js
```

The wizard will prompt for:
- Company Name (required)
- Short Name for menus
- Menu Icon emoji
- **Industry selection** (auto-suggested based on emoji)
- Company Email (optional)
- Company Phone (optional)

### Option 2: Manual Setup

1. Edit `core/config-module.gs` - Replace placeholders:
   - `{{COMPANY_NAME}}` - Your company name
   - `{{SHORT_NAME}}` - Short name for menus
   - `{{COMPANY_EMAIL}}` - Contact email
   - `{{COMPANY_PHONE}}` - Contact phone
   - `{{MENU_ICON}}` - Menu emoji (e.g., 🏢)

2. Login to clasp:
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

3. Create and deploy:
   ```bash
   clasp create --type sheets --title "Company Workspace Tools"
   clasp push
   clasp open
   ```

4. Authorize the script when prompted

5. Open a Google Sheet and refresh to see the menu

## Industry Templates

During setup, choose an industry template that matches your business:

| Industry | Emoji | Folder Structure | Gmail Labels | Auto-Label Rules |
|----------|-------|------------------|--------------|------------------|
| Consulting / Agency | 💼 | Clients, Projects, Deliverables, Proposals | Clients, Projects, Invoices, Proposals | Client domains, Project tools, Accounting |
| Real Estate | 🏠 | Listings, Transactions, Clients, Marketing | Buyers, Sellers, Listings, Transactions | MLS, Title, Lenders, Clients |
| Construction / Trades | 🔧 | Jobs, Estimates, Permits, Subcontractors | Jobs, Permits, Suppliers, Subcontractors | Permit offices, Suppliers, Clients |
| Healthcare / Medical | 🏥 | Patients, Insurance, Clinical, Administrative | Patients, Insurance, Labs, Referrals | Insurance, Labs, Pharmacy, Referrals |
| E-Commerce / Retail | 🛒 | Products, Orders, Customers, Marketing | Orders, Customers, Suppliers, Returns | Platforms, Shipping, Customers, Support |
| Technology / SaaS | 💻 | Products, Customers, Development, Support | Customers, Support, Development, Sales | Customer domains, Platforms, Dev tools |
| Analytics / Data | 📊 | Projects, Datasets, Reports, Clients | Clients, Projects, Reports, Data Sources | Client domains, BI tools, Data sources |
| Fleet / Automotive | 🚗 | Vehicles, Maintenance, Drivers, Compliance | Fleet, Maintenance, Insurance, Compliance | Fuel, Maintenance, Insurance providers |
| Strategy / Premium | ⭐ | Engagements, Research, Deliverables, Clients | Clients, Engagements, Research, Reports | Client domains, Research sources |
| General Business | 🏢 | Clients, Projects, Finance, Operations | Clients, Projects, Vendors, Internal | Client domains, Vendors, Internal |

Each template includes:
- **Folder structure** - Organized Google Drive hierarchy for your industry
- **Gmail labels** - Color-coded label system for email organization
- **Auto-label rules** - Automatic email labeling by sender domain/subject

## Configuration Placeholders

| Placeholder | Description |
|-------------|-------------|
| `{{COMPANY_NAME}}` | Full company name |
| `{{SHORT_NAME}}` | Short name for menus |
| `{{MENU_ICON}}` | Emoji icon for menu |
| `{{COMPANY_EMAIL}}` | Company contact email |
| `{{COMPANY_PHONE}}` | Company phone number |
| `{{INDUSTRY_KEY}}` | Selected industry identifier |
| `{{INDUSTRY_NAME}}` | Industry display name |
| `{{FOLDER_STRUCTURE}}` | JSON array of folder structure |
| `{{GMAIL_LABELS}}` | JSON array of Gmail labels |
| `{{AUTO_LABEL_RULES}}` | JSON array of auto-label rules |

## File Structure

```
g-stack-workspace-command/
├── core/
│   ├── WorkspaceTools.gs       # Main menu, sidebar, dashboard controller
│   ├── config-module.gs        # Company configuration (edit this!)
│   ├── gmail-automation.gs     # Daily digest, lead capture
│   ├── gmail-label-creator.gs  # Label creation and auto-labeling
│   ├── gmail-module.gs         # Gmail dispatcher helpers
│   ├── drive-inventory.gs      # Drive scanning and analysis
│   ├── drive-folder-setup.gs   # Folder structure creation
│   ├── drive-module.gs         # Drive dispatcher helpers
│   ├── logger.gs               # Logging utilities
│   └── runner.gs               # Action dispatcher
├── ui/
│   ├── Dashboard.html          # Interactive dashboard UI
│   ├── Sidebar.html            # Sidebar UI
│   ├── Help.html               # In-app help documentation
│   ├── UserManual.html         # Detailed user manual dialog
│   └── presentation.html       # Presentation view
├── tests/
│   ├── TestRunner.gs           # Test execution framework
│   └── Tests.gs                # Unit test definitions
├── docs/
│   ├── README.md               # This file
│   └── USER-MANUAL.md          # Markdown user manual
├── setup/
│   └── setup.js                # Automated setup script
├── config/
│   └── config.json             # Theme/branding config
├── appsscript.json             # Apps Script manifest (OAuth scopes)
├── .clasp.json.template        # clasp configuration template
├── LICENSE                     # MIT
├── g-stack-workspace.png       # Logo
└── README.md                   # Top-level overview
```

## OAuth Scopes Required

- `spreadsheets` - Read/write spreadsheet data
- `drive` - Create folders and scan files
- `gmail.modify` - Archive and label emails
- `gmail.labels` - Create and manage labels
- `script.send_mail` - Send digest emails
- `script.container.ui` - Display menus and dialogs

## Sheets Created

When you run "Initialize All Tools", these sheets are created:

| Sheet Name | Description |
|------------|-------------|
| Drive Inventory | File catalog with metadata |
| Folder Structure | Folder hierarchy |
| Potential Duplicates | Duplicate file report |
| Move Recommendations | Suggested file moves |
| Summary | Drive scan summary |
| Gmail Labels | Label list and status |
| Auto-Label Rules | Email labeling rules |
| Captured Leads | Leads extracted from emails |
| Email Statistics | Email volume tracking |
| Response Templates | Email templates |

## Triggers (Automated Tasks)

After running "Create All Triggers":

| Trigger | Time | Function |
|---------|------|----------|
| Daily Digest | 7 AM | Sends email summary |
| Lead Capture | Every 2 hours | Captures leads from portal emails |
| Auto-Label | Every 15 min | Labels new emails |

## For Real Estate

If you need real estate features (Property Listings, Transaction Tracker, Calendar Sync), use the **Realty Module** instead. This module is designed for general business workspace setup.

## Support

For help or feature requests:
- Email: jacob@truenorthstrategyops.com
- Create an issue in this repository

---

**TNDS Workspace Tools** - Built for productivity
