# Workspace Tools - User Manual

**Version:** 1.0
**Company:** {{COMPANY_NAME}}
**Industry:** {{INDUSTRY_NAME}}
**Support:** jacob@truenorthstrategyops.com

---

## Table of Contents

1. [Initial Deployment](#initial-deployment)
2. [Getting Started](#getting-started)
3. [Available Industry Templates](#available-industry-templates)
4. [Drive Folder Tools](#drive-folder-tools)
5. [Gmail Label Tools](#gmail-label-tools)
6. [Troubleshooting](#troubleshooting)

---

## Initial Deployment

### Running the Setup Wizard

When you run `node setup.js` in the module folder, the wizard prompts you to configure:

#### 1. Company Information

```
Company Name: Your Company Name
Short Name (for menus): YourCo
Company Email: <company-email-placeholder>
Company Phone: (555) 123-4567
```

#### 2. Menu Emoji Selection

Choose an emoji that appears in the Google Sheets menu:

```
Select menu emoji:
  1) 🏢 Building (General)
  2) 🏠 House (Real Estate)
  3) 💼 Briefcase (Consulting)
  4) 🔧 Wrench (Construction)
  5) 🏥 Hospital (Healthcare)
  6) 🛒 Cart (E-Commerce)
  7) 💻 Computer (Technology)
  8) 📊 Chart (Analytics)
  9) 🚗 Car (Fleet)
  10) ⭐ Star (Premium)
```

#### 3. Industry Template Selection

Your industry determines the folder structure, Gmail labels, and auto-label rules:

```
Select your industry:
  1) 💼 Consulting / Agency
  2) 🏠 Real Estate
  3) 🔧 Construction / Trades
  4) 🏥 Healthcare / Medical
  5) 🛒 E-Commerce / Retail
  6) 💻 Technology / SaaS
  7) 💰 Finance / Accounting
  8) ⚖️ Legal / Law Firm
  9) 📚 Education / Training
  10) 🚗 Fleet / Automotive
  11) 📊 Analytics / Data
  12) 🏢 General Business
```

#### What Gets Created

After you confirm, the wizard:
1. Creates a customized output folder with your company name
2. Replaces all placeholders in template files
3. Runs `clasp create` to set up a new Google Sheet
4. Runs `clasp push` to deploy the Apps Script code
5. Opens the Google Sheet in your browser

**To change settings:** Re-run `node setup.js` to create a new deployment with different options.

---

## Getting Started

### First-Time Setup (After Deployment)

1. **Open the Google Sheet** that was created during setup
2. **Refresh the page** (Ctrl+R or Cmd+R) to load the custom menu
3. **Click the menu** in the toolbar (look for the {{MENU_ICON}} icon)
4. Go to **Setup & Initialize > Initialize All Sheets**
5. **Grant permissions** when prompted

### Quick Start Actions

| Action | Menu Path |
|--------|-----------|
| Create Drive folders | Drive Tools > Create Folder Structure |
| Create Gmail labels | Gmail Tools > Create Gmail Labels |
| Open quick dashboard | Open Sidebar |
| View this help | Help & About > User Manual |

---

## Available Industry Templates

When setting up Workspace Tools, you can choose from these industry templates:

| Industry | Emoji | Best For |
|----------|-------|----------|
| Real Estate | 🏠 | Agents, brokerages, property managers |
| Consulting / Agency | 💼 | Consultants, agencies, professional services |
| Construction / Trades | 🔧 | Contractors, construction, trades businesses |
| Healthcare / Medical | 🏥 | Medical practices, healthcare providers |
| E-Commerce / Retail | 🛒 | Online stores, retail businesses |
| General Business | 🏢 | General purpose for any business |
| Analytics / Data | 📊 | Data analysts, BI consultants |
| Fleet / Automotive | 🚗 | Fleet management, automotive businesses |
| Strategy / Premium | ⭐ | Strategy consulting, premium services |

Each template includes:
- **Folder structure** for Google Drive
- **Gmail labels** for email organization
- **Auto-label rules** for incoming emails

---

## Drive Folder Tools

### Create Folder Structure

Creates an organized folder hierarchy in your Google Drive based on your industry template.

**Your folder structure ({{INDUSTRY_NAME}}):**

{{FOLDER_STRUCTURE}}

**How to use:**
1. Menu > Drive Tools > Create Folder Structure
2. Confirm when prompted
3. Folder link will be displayed when complete
4. Share the root folder with team members

### Scan My Drive

Catalogs all files in your Google Drive with metadata.

**Creates sheets:**
- **Drive Inventory** - List of all files with details
- **Folder Structure** - All folders and sizes
- **Summary** - Overview statistics

**Tips:**
- Scans up to 1,000 files (to prevent timeout)
- Run periodically to keep inventory updated
- Use "Find Duplicates" to identify duplicate files

---

## Gmail Label Tools

### Create Gmail Labels

Creates color-coded labels for email organization based on your industry template.

**Your label structure ({{INDUSTRY_NAME}}):**

The labels are customized for your business type and include parent labels and sub-labels for detailed organization.

**How to use:**
1. Menu > Gmail Tools > Create Gmail Labels
2. Labels appear immediately in Gmail
3. Use "Delete All Created Labels" to remove if needed

### Auto-Label Rules

Automatically labels incoming emails based on sender and subject patterns.

**How to customize:**
1. Go to the **Auto-Label Rules** sheet
2. Edit existing rules or add new rows
3. Set **Active** checkbox to enable/disable rules

**Columns:**
- **Label** - The Gmail label to apply
- **From Contains** - Email domains to match (comma-separated)
- **Subject Contains** - Keywords in subject to match
- **Active** - Checkbox to enable the rule

### Label Migration

Migrate from old Gmail labels to the new organized structure.

**Steps:**
1. Gmail Tools > Label Migration > Scan Labels
2. Review suggestions in the Label Migration sheet
3. Check "Move Emails?" for labels to migrate
4. Gmail Tools > Label Migration > Preview Migration
5. Gmail Tools > Label Migration > Execute Migration

---

## Troubleshooting

### Menu Not Showing

- **Solution:** Refresh the page (Ctrl+R or Cmd+R)
- **Alternative:** Close and reopen the spreadsheet

### Permission Errors

- **Solution:** Re-authorize when prompted
- **Check:** Extensions > Apps Script > Triggers

### Functions Timing Out

- **Solution:** Try running on fewer items
- **For Drive scans:** Limit is 1,000 files per scan

### Labels Not Appearing in Gmail

- **Check:** Refresh Gmail (F5)
- **Verify:** Menu > Gmail Tools > List Existing Labels

### Need Help?

Contact: **jacob@truenorthstrategyops.com**

---

## About

**Workspace Tools** by True North Data Strategies

Workspace automation tools for {{INDUSTRY_NAME}} businesses including:
- Google Drive folder organization
- Gmail label management and auto-labeling
- Email analytics and lead capture
- File inventory and cleanup recommendations

*This toolkit was customized for {{COMPANY_NAME}}.*
