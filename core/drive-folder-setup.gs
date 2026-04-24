/**
 * Drive Folder Setup
 * TNDS Workspace Automation Kit
 *
 * Creates organized folder structure for {{INDUSTRY_NAME}} business
 * Run once, then share root folder with team members
 */

// ============================================
// Configuration
// ============================================
const FOLDER_CONFIG = {
  // Uses COMPANY_CONFIG.NAME if available, otherwise placeholder
  ROOT_NAME: typeof COMPANY_CONFIG !== 'undefined' ? COMPANY_CONFIG.ROOT_FOLDER_NAME : "{{COMPANY_NAME}}",
  INDUSTRY: "{{INDUSTRY_KEY}}",
  STRUCTURE: {{FOLDER_STRUCTURE}}
};

// ============================================
// Main Function - Run This
// ============================================
/**
 * Creates the complete folder structure
 * Run this function from the Apps Script editor
 */
function createFolderStructure() {
  try {
    Logger.log('Starting folder structure creation...');

    // Check if root folder already exists
    const existingFolders = DriveApp.getFoldersByName(FOLDER_CONFIG.ROOT_NAME);
    if (existingFolders.hasNext()) {
      const existing = existingFolders.next();
      Logger.log(`Root folder already exists: ${existing.getUrl()}`);

      // Ask if they want to use existing or create new
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Folder Exists',
        `"${FOLDER_CONFIG.ROOT_NAME}" already exists.\n\nDo you want to add missing subfolders to it?`,
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        createSubfolders(existing);
        showCompletionDialog(existing);
      }
      return;
    }

    // Create root folder
    const rootFolder = DriveApp.createFolder(FOLDER_CONFIG.ROOT_NAME);
    Logger.log(`Created root folder: ${rootFolder.getName()}`);

    // Create subfolders
    createSubfolders(rootFolder);

    // Show completion dialog
    showCompletionDialog(rootFolder);

  } catch (error) {
    Logger.log(`ERROR: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error creating folders: ${error.message}`);
  }
}

/**
 * Creates all subfolders within the root folder
 * @param {Folder} rootFolder - The root folder to create structure in
 */
function createSubfolders(rootFolder) {
  const structure = FOLDER_CONFIG.STRUCTURE;

  for (const mainFolder in structure) {
    // Create or get main subfolder
    let mainFolderObj = getOrCreateFolder(rootFolder, mainFolder);
    Logger.log(`  Created/Found: ${mainFolder}`);

    // Create children
    const children = structure[mainFolder];
    children.forEach(childName => {
      getOrCreateFolder(mainFolderObj, childName);
      Logger.log(`    Created/Found: ${childName}`);
    });
  }
}

/**
 * Gets existing folder or creates new one
 * @param {Folder} parent - Parent folder
 * @param {string} name - Folder name to find/create
 * @returns {Folder} - The folder object
 */
function getOrCreateFolder(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parent.createFolder(name);
}

/**
 * Shows completion dialog with folder link
 * @param {Folder} rootFolder - The created root folder
 */
function showCompletionDialog(rootFolder) {
  const folderUrl = rootFolder.getUrl();
  const folderId = rootFolder.getId();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { color: #1a73e8; }
      .success { color: #34a853; font-weight: bold; }
      .url-box {
        background: #f1f3f4;
        padding: 10px;
        border-radius: 4px;
        word-break: break-all;
        margin: 10px 0;
      }
      .btn {
        background: #1a73e8;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        margin-top: 10px;
      }
      .btn:hover { background: #1557b0; }
      .note { color: #666; font-size: 12px; margin-top: 15px; }
    </style>

    <h2>Folder Structure Created!</h2>
    <p class="success">All folders have been created successfully.</p>

    <h3>Folder URL:</h3>
    <div class="url-box">${folderUrl}</div>

    <h3>Folder ID:</h3>
    <div class="url-box">${folderId}</div>

    <a href="${folderUrl}" target="_blank">
      <button class="btn">Open in Drive</button>
    </a>

    <p class="note">
      <strong>Next step:</strong> Right-click the root folder in Drive and select
      "Share" to give access to team members or other accounts.
    </p>
  `)
  .setWidth(450)
  .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Folder Setup Complete');
}

// ============================================
// Utility Functions
// ============================================

/**
 * Lists all folders in the created structure
 * Useful for verification
 */
function listFolderStructure() {
  const folders = DriveApp.getFoldersByName(FOLDER_CONFIG.ROOT_NAME);

  if (!folders.hasNext()) {
    Logger.log('Root folder not found. Run createFolderStructure() first.');
    return;
  }

  const root = folders.next();
  Logger.log(`\n=== ${root.getName()} ===`);
  Logger.log(`URL: ${root.getUrl()}`);

  listFoldersRecursive(root, 1);
}

/**
 * Recursively lists folders
 * @param {Folder} folder - Folder to list
 * @param {number} depth - Current depth for indentation
 */
function listFoldersRecursive(folder, depth) {
  const indent = '  '.repeat(depth);
  const subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    Logger.log(`${indent}├── ${sub.getName()}`);
    listFoldersRecursive(sub, depth + 1);
  }
}

/**
 * Shares the root folder with another email address
 * @param {string} email - Email address to share with
 * @param {string} role - 'edit' or 'view'
 */
function shareRootFolder(email, role) {
  const folders = DriveApp.getFoldersByName(FOLDER_CONFIG.ROOT_NAME);

  if (!folders.hasNext()) {
    Logger.log('Root folder not found. Run createFolderStructure() first.');
    return;
  }

  const root = folders.next();

  if (role === 'edit') {
    root.addEditor(email);
  } else {
    root.addViewer(email);
  }

  Logger.log(`Shared "${root.getName()}" with ${email} (${role} access)`);
}

/**
 * Quick share with True North
 * Modify the email as needed
 */
function shareWithTrueNorth() {
  // Update this email to your truenorth account
  const truenorthEmail = 'jacob@truenorthstrategyops.com';
  shareRootFolder(truenorthEmail, 'edit');
}

// ============================================
// Menu Setup (disabled - using unified menu in main Tools file)
// ============================================
function _onOpenDriveSetup() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Drive Setup')
      .addItem('Create Folder Structure', 'createFolderStructure')
      .addItem('List Folder Structure', 'listFolderStructure')
      .addSeparator()
      .addItem('Share with True North', 'shareWithTrueNorth')
      .addToUi();
  } catch (e) {
    // Not in a spreadsheet context - that's fine
  }
}
