/**
 * Drive Inventory & Organization Assistant
 * TNDS Workspace Automation Kit
 *
 * Features:
 * - Scans entire Drive or specific folders
 * - Catalogs all files and folders with metadata
 * - Identifies duplicates, large files, old files
 * - Recommends organization based on file types
 * - Generates move recommendations to new folder structure
 */

// ============================================
// Configuration
// ============================================
const INVENTORY_CONFIG = {
  SHEETS: {
    INVENTORY: 'Drive Inventory',
    FOLDERS: 'Folder Structure',
    DUPLICATES: 'Potential Duplicates',
    RECOMMENDATIONS: 'Move Recommendations',
    SUMMARY: 'Summary'
  },
  // Target folder structure for recommendations
  TARGET_FOLDERS: {
    'Listings': ['listing', 'mls', 'property', 'home', 'house', 'condo', 'real estate'],
    'Clients': ['client', 'buyer', 'seller', 'customer', 'contact'],
    'Transactions': ['contract', 'agreement', 'closing', 'escrow', 'title', 'deed', 'transaction'],
    'Marketing': ['flyer', 'brochure', 'marketing', 'social', 'ad', 'promo', 'campaign'],
    'Operations': ['policy', 'procedure', 'training', 'license', 'insurance', 'vendor']
  },
  // File type categories
  FILE_CATEGORIES: {
    'Documents': ['application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.google-apps.document'],
    'Spreadsheets': ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.google-apps.spreadsheet'],
    'Presentations': ['application/vnd.ms-powerpoint', 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'application/vnd.google-apps.presentation'],
    'Images': ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'image/heic'],
    'Videos': ['video/mp4', 'video/quicktime', 'video/x-msvideo'],
    'Audio': ['audio/mpeg', 'audio/wav', 'audio/x-m4a']
  },
  MAX_FILES: 1000,  // Limit to prevent timeout

  // Developer project file detection
  // Files that commonly appear in every project and should NOT be flagged as duplicates
  PROJECT_FILES: [
    'index.html', 'index.js', 'index.ts', 'index.tsx', 'index.css',
    'package.json', 'package-lock.json', 'yarn.lock', 'pnpm-lock.yaml',
    'tsconfig.json', 'jsconfig.json', 'vite.config.js', 'vite.config.ts',
    'webpack.config.js', 'rollup.config.js', 'babel.config.js',
    '.gitignore', '.eslintrc', '.eslintrc.js', '.eslintrc.json', '.prettierrc',
    'README.md', 'readme.md', 'LICENSE', 'CHANGELOG.md',
    'Dockerfile', 'docker-compose.yml', 'docker-compose.yaml',
    '.env', '.env.example', '.env.local', '.env.development', '.env.production',
    'app.js', 'main.js', 'main.ts', 'App.js', 'App.tsx', 'App.vue',
    'styles.css', 'style.css', 'global.css', 'globals.css',
    'manifest.json', 'robots.txt', 'sitemap.xml', 'favicon.ico',
    'next.config.js', 'nuxt.config.js', 'gatsby-config.js', 'astro.config.mjs',
    'tailwind.config.js', 'tailwind.config.ts', 'postcss.config.js',
    'jest.config.js', 'vitest.config.ts', 'playwright.config.ts',
    'appsscript.json', 'clasp.json', '.clasp.json'
  ],

  // Folders that indicate a project root
  PROJECT_ROOT_INDICATORS: [
    'node_modules', '.git', 'package.json', 'Cargo.toml', 'go.mod',
    'requirements.txt', 'Pipfile', 'pyproject.toml', 'setup.py',
    'Gemfile', 'composer.json', 'pom.xml', 'build.gradle'
  ]
};

// ============================================
// Menu Setup (disabled - using unified menu in main Tools file)
// ============================================
function _onOpenDriveInventory() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📁 Drive Inventory')
    .addItem('🔍 Scan My Drive', 'scanMyDrive')
    .addItem('📂 Scan Specific Folder', 'scanSpecificFolder')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Analysis')
      .addItem('Find Duplicates', 'findDuplicates')
      .addItem('Find Large Files (>10MB)', 'findLargeFiles')
      .addItem('Find Old Files (>1 year)', 'findOldFiles')
      .addItem('Generate Move Recommendations', 'generateRecommendations'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('Initialize Sheets', 'initializeInventorySheets')
      .addItem('Clear Inventory', 'clearInventory'))
    .addToUi();
}

// ============================================
// Sheet Initialization
// ============================================
function initializeInventorySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Inventory sheet
  let invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
  if (!invSheet) {
    invSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.INVENTORY);
  }
  invSheet.clear();
  const invHeaders = ['File Name', 'File Type', 'MIME Type', 'Size (KB)', 'Created', 'Last Updated', 'Owner', 'Path', 'File ID', 'URL', 'Category', 'Suggested Folder'];
  invSheet.getRange(1, 1, 1, invHeaders.length).setValues([invHeaders]);
  invSheet.getRange(1, 1, 1, invHeaders.length)
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  invSheet.setFrozenRows(1);
  invSheet.setColumnWidth(1, 250);
  invSheet.setColumnWidth(8, 300);

  // Folders sheet
  let foldersSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.FOLDERS);
  if (!foldersSheet) {
    foldersSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.FOLDERS);
  }
  foldersSheet.clear();
  const folderHeaders = ['Folder Name', 'Path', 'File Count', 'Total Size (MB)', 'Folder ID', 'URL'];
  foldersSheet.getRange(1, 1, 1, folderHeaders.length).setValues([folderHeaders]);
  foldersSheet.getRange(1, 1, 1, folderHeaders.length)
    .setBackground('#2e7d32')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  foldersSheet.setFrozenRows(1);

  // Duplicates sheet
  let dupSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.DUPLICATES);
  if (!dupSheet) {
    dupSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.DUPLICATES);
  }
  dupSheet.clear();
  const dupHeaders = ['File Name', 'Size (KB)', 'Location 1', 'Location 2', 'Action'];
  dupSheet.getRange(1, 1, 1, dupHeaders.length).setValues([dupHeaders]);
  dupSheet.getRange(1, 1, 1, dupHeaders.length)
    .setBackground('#e65100')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  dupSheet.setFrozenRows(1);

  // Recommendations sheet
  let recSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.RECOMMENDATIONS);
  if (!recSheet) {
    recSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.RECOMMENDATIONS);
  }
  recSheet.clear();
  const recHeaders = ['File Name', 'Current Location', 'Recommended Folder', 'Reason', 'Confidence', 'Move?', 'File ID'];
  recSheet.getRange(1, 1, 1, recHeaders.length).setValues([recHeaders]);
  recSheet.getRange(1, 1, 1, recHeaders.length)
    .setBackground('#6a1b9a')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  recSheet.setFrozenRows(1);

  // Checkbox for Move column
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  recSheet.getRange(2, 6, 500, 1).setDataValidation(checkboxRule);

  // Summary sheet
  let sumSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.SUMMARY);
  if (!sumSheet) {
    sumSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.SUMMARY);
  }
  sumSheet.clear();

  SpreadsheetApp.getUi().alert('✅ Inventory sheets initialized!');
}

function clearInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.values(INVENTORY_CONFIG.SHEETS).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
    }
  });

  SpreadsheetApp.getUi().alert('Inventory cleared.');
}

// ============================================
// Scanning Functions
// ============================================
function scanMyDrive() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Scan My Drive',
    'This will scan your entire Drive (up to ' + INVENTORY_CONFIG.MAX_FILES + ' files).\n\nThis may take a few minutes. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
  let foldersSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.FOLDERS);

  if (!invSheet || !foldersSheet) {
    initializeInventorySheets();
    invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
    foldersSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.FOLDERS);
  }

  // Clear existing data
  if (invSheet.getLastRow() > 1) {
    invSheet.getRange(2, 1, invSheet.getLastRow() - 1, invSheet.getLastColumn()).clear();
  }
  if (foldersSheet.getLastRow() > 1) {
    foldersSheet.getRange(2, 1, foldersSheet.getLastRow() - 1, foldersSheet.getLastColumn()).clear();
  }

  const rootFolder = DriveApp.getRootFolder();
  const results = { files: [], folders: [], count: 0 };

  ui.alert('Scanning started. Please wait...');

  scanFolder(rootFolder, 'My Drive', results);

  // Write results
  if (results.files.length > 0) {
    invSheet.getRange(2, 1, results.files.length, results.files[0].length).setValues(results.files);
  }
  if (results.folders.length > 0) {
    foldersSheet.getRange(2, 1, results.folders.length, results.folders[0].length).setValues(results.folders);
  }

  // Generate summary
  generateSummary(results);

  // Build completion message with error info if any
  let message = '✅ Scan complete!\n\nFiles found: ' + results.files.length + '\nFolders found: ' + results.folders.length;
  if (results.skipped && results.skipped > 0) {
    message += '\nSkipped files: ' + results.skipped;
  }
  if (results.errors && results.errors.length > 0) {
    message += '\nErrors encountered: ' + results.errors.length;
    message += '\n\n(Some files/folders may be inaccessible due to permissions)';
    // Log errors for debugging
    console.log('Scan errors:', JSON.stringify(results.errors.slice(0, 10)));
  }
  ui.alert(message);
}

function scanSpecificFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Scan Specific Folder',
    'Enter the folder URL or folder ID:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  let folderId;

  // Extract folder ID from URL if needed
  if (input.includes('folders/')) {
    folderId = input.split('folders/')[1].split('?')[0];
  } else {
    folderId = input;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
    let foldersSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.FOLDERS);

    if (!invSheet || !foldersSheet) {
      initializeInventorySheets();
      invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
      foldersSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.FOLDERS);
    }

    const results = { files: [], folders: [], count: 0 };

    scanFolder(folder, folder.getName(), results);

    // Append results
    if (results.files.length > 0) {
      const lastRow = invSheet.getLastRow();
      invSheet.getRange(lastRow + 1, 1, results.files.length, results.files[0].length).setValues(results.files);
    }
    if (results.folders.length > 0) {
      const lastRow = foldersSheet.getLastRow();
      foldersSheet.getRange(lastRow + 1, 1, results.folders.length, results.folders[0].length).setValues(results.folders);
    }

    ui.alert('✅ Folder scanned!\n\nFiles found: ' + results.files.length + '\nFolders found: ' + results.folders.length);

  } catch (e) {
    ui.alert('Error: Could not access folder.\n\n' + e.message);
  }
}

function scanFolder(folder, path, results) {
  if (results.count >= INVENTORY_CONFIG.MAX_FILES) return;

  // Initialize error tracking if not present
  if (!results.errors) results.errors = [];
  if (!results.skipped) results.skipped = 0;

  let folderSize = 0;
  let folderFileCount = 0;
  let folderName = 'Unknown';
  let folderId = '';
  let folderUrl = '';

  try {
    folderName = folder.getName();
    folderId = folder.getId();
    folderUrl = folder.getUrl();
  } catch (e) {
    results.errors.push({ path: path, error: 'Cannot access folder metadata: ' + e.message });
    return;  // Skip this folder entirely
  }

  // Scan files in this folder with error handling
  try {
    const files = folder.getFiles();

    while (files.hasNext() && results.count < INVENTORY_CONFIG.MAX_FILES) {
      try {
        const file = files.next();
        const size = file.getSize();
        folderSize += size;
        folderFileCount++;
        results.count++;

        const mimeType = file.getMimeType();
        const category = getFileCategory(mimeType);
        const suggestedFolder = getSuggestedFolder(file.getName(), mimeType);

        // Get owner safely - some files may not have accessible owner
        let ownerEmail = 'Unknown';
        try {
          const owner = file.getOwner();
          if (owner) ownerEmail = owner.getEmail();
        } catch (ownerErr) {
          ownerEmail = 'Inaccessible';
        }

        results.files.push([
          file.getName(),
          getFileExtension(file.getName()),
          mimeType,
          Math.round(size / 1024),  // KB
          file.getDateCreated(),
          file.getLastUpdated(),
          ownerEmail,
          path,
          file.getId(),
          file.getUrl(),
          category,
          suggestedFolder
        ]);

        // Throttle to avoid rate limits (pause every 50 files)
        if (results.count % 50 === 0) {
          Utilities.sleep(100);  // 100ms pause
        }

      } catch (fileErr) {
        results.skipped++;
        results.errors.push({ path: path, error: 'File error: ' + fileErr.message });
        // Continue with next file instead of crashing
      }
    }
  } catch (filesErr) {
    results.errors.push({ path: path, error: 'Cannot list files: ' + filesErr.message });
    // Continue to record folder and scan subfolders if possible
  }

  // Record this folder
  results.folders.push([
    folderName,
    path,
    folderFileCount,
    Math.round(folderSize / (1024 * 1024) * 100) / 100,  // MB
    folderId,
    folderUrl
  ]);

  // Recursively scan subfolders with error handling
  try {
    const subfolders = folder.getFolders();
    while (subfolders.hasNext() && results.count < INVENTORY_CONFIG.MAX_FILES) {
      try {
        const subfolder = subfolders.next();
        scanFolder(subfolder, path + '/' + subfolder.getName(), results);

        // Throttle between folders to avoid rate limits
        Utilities.sleep(50);  // 50ms pause
      } catch (subErr) {
        results.errors.push({ path: path, error: 'Subfolder error: ' + subErr.message });
        // Continue with next subfolder
      }
    }
  } catch (subfoldersErr) {
    results.errors.push({ path: path, error: 'Cannot list subfolders: ' + subfoldersErr.message });
  }
}

// ============================================
// Helper Functions
// ============================================
function getFileExtension(filename) {
  const parts = filename.split('.');
  return parts.length > 1 ? parts.pop().toUpperCase() : 'Unknown';
}

function getFileCategory(mimeType) {
  for (const [category, types] of Object.entries(INVENTORY_CONFIG.FILE_CATEGORIES)) {
    if (types.includes(mimeType)) return category;
  }
  if (mimeType.startsWith('application/vnd.google-apps')) {
    return 'Google ' + mimeType.split('.').pop();
  }
  return 'Other';
}

function getSuggestedFolder(filename, mimeType) {
  const lowerName = filename.toLowerCase();

  for (const [folder, keywords] of Object.entries(INVENTORY_CONFIG.TARGET_FOLDERS)) {
    for (const keyword of keywords) {
      if (lowerName.includes(keyword)) {
        return folder;
      }
    }
  }

  // Suggest based on file type
  if (mimeType.startsWith('image/')) return 'Listings/Photos & Media';
  if (mimeType.startsWith('video/')) return 'Marketing/Social Media';

  return 'Review Needed';
}

// ============================================
// Analysis Functions
// ============================================
function findDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
  let dupSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.DUPLICATES);

  if (!invSheet || invSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No inventory data. Run a scan first.');
    return;
  }

  // Create or update duplicates sheet with new headers
  if (!dupSheet) {
    dupSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.DUPLICATES);
  }
  dupSheet.clear();
  const headers = ['File Name', 'Size (KB)', 'Location 1', 'Location 2', 'Type', 'Safe to Delete?', 'Action'];
  dupSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  dupSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#e65100')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  dupSheet.setFrozenRows(1);
  dupSheet.setColumnWidth(1, 200);
  dupSheet.setColumnWidth(3, 250);
  dupSheet.setColumnWidth(4, 250);

  const data = invSheet.getDataRange().getValues();
  const fileMap = {};
  const duplicates = [];

  // Check if a file is a common project file
  const isProjectFile = (filename) => {
    const lower = filename.toLowerCase();
    return INVENTORY_CONFIG.PROJECT_FILES.some(pf => lower === pf.toLowerCase());
  };

  // Extract project root from path (e.g., "My Drive/Projects/my-app/src" -> "My Drive/Projects/my-app")
  const getProjectRoot = (path) => {
    const parts = path.split('/');
    // Find the deepest folder that might be a project root
    // For simplicity, assume project is 2-3 levels deep from My Drive
    if (parts.length >= 3) {
      return parts.slice(0, 3).join('/');
    }
    return path;
  };

  // Group by filename and size
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0] + '_' + data[i][3];  // name_size
    if (!fileMap[key]) {
      fileMap[key] = [];
    }
    fileMap[key].push({
      name: data[i][0],
      size: data[i][3],
      path: data[i][7],
      projectRoot: getProjectRoot(data[i][7])
    });
  }

  // Find duplicates with smart categorization
  for (const [key, files] of Object.entries(fileMap)) {
    if (files.length > 1) {
      const filename = files[0].name;
      const isProjFile = isProjectFile(filename);

      // Check if files are in different project roots
      const uniqueRoots = new Set(files.map(f => f.projectRoot));
      const inDifferentProjects = uniqueRoots.size > 1;

      // Determine duplicate type and safety
      let dupType, safeToDelete;
      if (isProjFile && inDifferentProjects) {
        dupType = 'Project File';
        safeToDelete = 'NO - Different projects';
      } else if (isProjFile) {
        dupType = 'Project File';
        safeToDelete = 'Review carefully';
      } else if (inDifferentProjects) {
        dupType = 'Different folders';
        safeToDelete = 'Review - may be intentional';
      } else {
        dupType = 'True Duplicate';
        safeToDelete = 'Likely safe';
      }

      // For files with more than 2 duplicates, show all locations
      const allPaths = files.map(f => f.path);

      duplicates.push([
        filename,
        files[0].size,
        allPaths[0],
        allPaths.length > 2 ? allPaths.slice(1).join(' | ') : (allPaths[1] || ''),
        dupType,
        safeToDelete,
        ''  // Action column for user
      ]);
    }
  }

  // Sort: True Duplicates first, then others
  duplicates.sort((a, b) => {
    if (a[4] === 'True Duplicate' && b[4] !== 'True Duplicate') return -1;
    if (a[4] !== 'True Duplicate' && b[4] === 'True Duplicate') return 1;
    return 0;
  });

  // Write data
  if (duplicates.length > 0) {
    dupSheet.getRange(2, 1, duplicates.length, duplicates[0].length).setValues(duplicates);

    // Color code by type
    for (let i = 0; i < duplicates.length; i++) {
      const rowNum = i + 2;
      if (duplicates[i][4] === 'Project File') {
        dupSheet.getRange(rowNum, 1, 1, headers.length).setBackground('#e3f2fd');  // Light blue
      } else if (duplicates[i][4] === 'True Duplicate') {
        dupSheet.getRange(rowNum, 1, 1, headers.length).setBackground('#ffebee');  // Light red
      }
    }
  }

  // Count by type
  const trueDups = duplicates.filter(d => d[4] === 'True Duplicate').length;
  const projFiles = duplicates.filter(d => d[4] === 'Project File').length;
  const otherDups = duplicates.length - trueDups - projFiles;

  SpreadsheetApp.getUi().alert(
    `Duplicate Analysis Complete!\n\n` +
    `True Duplicates: ${trueDups} (likely safe to delete)\n` +
    `Project Files: ${projFiles} (probably intentional)\n` +
    `Other: ${otherDups}\n\n` +
    `Check the "${INVENTORY_CONFIG.SHEETS.DUPLICATES}" sheet.\n` +
    `Red rows = likely true duplicates\n` +
    `Blue rows = project files (probably keep)`
  );
}

function findLargeFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);

  if (!invSheet || invSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No inventory data. Run a scan first.');
    return;
  }

  const data = invSheet.getDataRange().getValues();
  let report = '📦 LARGE FILES (>10MB)\n\n';
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const sizeKB = data[i][3];
    if (sizeKB > 10240) {  // 10MB in KB
      count++;
      const sizeMB = Math.round(sizeKB / 1024);
      report += `• ${data[i][0]}\n  ${sizeMB} MB | ${data[i][7]}\n\n`;
    }
  }

  if (count === 0) {
    report += 'No files larger than 10MB found.';
  } else {
    report = `Found ${count} large files:\n\n` + report;
  }

  SpreadsheetApp.getUi().alert(report);
}

function findOldFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);

  if (!invSheet || invSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No inventory data. Run a scan first.');
    return;
  }

  const data = invSheet.getDataRange().getValues();
  const oneYearAgo = new Date();
  oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

  let report = '📅 OLD FILES (Not updated in 1+ year)\n\n';
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const lastUpdated = data[i][5];
    if (lastUpdated instanceof Date && lastUpdated < oneYearAgo) {
      count++;
      if (count <= 20) {  // Limit display
        report += `• ${data[i][0]}\n  Last updated: ${lastUpdated.toLocaleDateString()}\n  Location: ${data[i][7]}\n\n`;
      }
    }
  }

  if (count === 0) {
    report += 'No files older than 1 year found.';
  } else {
    report = `Found ${count} old files${count > 20 ? ' (showing first 20)' : ''}:\n\n` + report;
  }

  SpreadsheetApp.getUi().alert(report);
}

// ============================================
// Recommendations
// ============================================
function generateRecommendations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.INVENTORY);
  const recSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.RECOMMENDATIONS);

  if (!invSheet || invSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No inventory data. Run a scan first.');
    return;
  }

  const data = invSheet.getDataRange().getValues();
  const recommendations = [];

  for (let i = 1; i < data.length; i++) {
    const fileName = data[i][0];
    const currentPath = data[i][7];
    const suggestedFolder = data[i][11];
    const fileId = data[i][8];

    if (suggestedFolder && suggestedFolder !== 'Review Needed') {
      // Check if already in suggested location
      if (!currentPath.includes(suggestedFolder)) {
        const confidence = getConfidenceLevel(fileName, suggestedFolder);
        const reason = getRecommendationReason(fileName, suggestedFolder);

        recommendations.push([
          fileName,
          currentPath,
          suggestedFolder,
          reason,
          confidence,
          false,  // Move checkbox
          fileId
        ]);
      }
    }
  }

  // Sort by confidence (High first)
  recommendations.sort((a, b) => {
    const order = { 'High': 0, 'Medium': 1, 'Low': 2 };
    return order[a[4]] - order[b[4]];
  });

  // Clear and write
  if (recSheet.getLastRow() > 1) {
    recSheet.getRange(2, 1, recSheet.getLastRow() - 1, recSheet.getLastColumn()).clear();
  }

  if (recommendations.length > 0) {
    recSheet.getRange(2, 1, recommendations.length, recommendations[0].length).setValues(recommendations);

    // Add conditional formatting for confidence
    const rules = recSheet.getConditionalFormatRules();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('High')
        .setBackground('#c8e6c9')
        .setRanges([recSheet.getRange('E:E')])
        .build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Medium')
        .setBackground('#fff3e0')
        .setRanges([recSheet.getRange('E:E')])
        .build()
    );
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Low')
        .setBackground('#ffebee')
        .setRanges([recSheet.getRange('E:E')])
        .build()
    );
    recSheet.setConditionalFormatRules(rules);
  }

  SpreadsheetApp.getUi().alert('Generated ' + recommendations.length + ' move recommendations.\n\nCheck the "' + INVENTORY_CONFIG.SHEETS.RECOMMENDATIONS + '" sheet.\n\nReview and check the "Move?" column for files you want to move.');
}

function getConfidenceLevel(filename, suggestedFolder) {
  const lowerName = filename.toLowerCase();
  const keywords = INVENTORY_CONFIG.TARGET_FOLDERS[suggestedFolder] || [];

  let matches = 0;
  for (const keyword of keywords) {
    if (lowerName.includes(keyword)) matches++;
  }

  if (matches >= 2) return 'High';
  if (matches === 1) return 'Medium';
  return 'Low';
}

function getRecommendationReason(filename, suggestedFolder) {
  const lowerName = filename.toLowerCase();
  const keywords = INVENTORY_CONFIG.TARGET_FOLDERS[suggestedFolder] || [];

  const matched = keywords.filter(k => lowerName.includes(k));
  if (matched.length > 0) {
    return 'Contains keywords: ' + matched.join(', ');
  }
  return 'File type suggests this category';
}

// ============================================
// Summary Generation
// ============================================
function generateSummary(results) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sumSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.SUMMARY);

  if (!sumSheet) {
    sumSheet = ss.insertSheet(INVENTORY_CONFIG.SHEETS.SUMMARY);
  }
  sumSheet.clear();

  // Calculate stats
  let totalSize = 0;
  const byCategory = {};
  const byFolder = {};

  results.files.forEach(file => {
    totalSize += file[3];  // Size KB
    const category = file[10];
    const suggested = file[11];

    byCategory[category] = (byCategory[category] || 0) + 1;
    byFolder[suggested] = (byFolder[suggested] || 0) + 1;
  });

  // Write summary
  const summary = [
    ['📊 DRIVE INVENTORY SUMMARY', ''],
    ['', ''],
    ['Total Files Scanned', results.files.length],
    ['Total Folders', results.folders.length],
    ['Total Size', Math.round(totalSize / 1024) + ' MB'],
    ['', ''],
    ['📁 BY CATEGORY', ''],
  ];

  Object.entries(byCategory).sort((a, b) => b[1] - a[1]).forEach(([cat, count]) => {
    summary.push([cat, count]);
  });

  summary.push(['', '']);
  summary.push(['📂 SUGGESTED DESTINATIONS', '']);

  Object.entries(byFolder).sort((a, b) => b[1] - a[1]).forEach(([folder, count]) => {
    summary.push([folder, count]);
  });

  sumSheet.getRange(1, 1, summary.length, 2).setValues(summary);

  // Format
  sumSheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sumSheet.getRange('A7').setFontWeight('bold');
  sumSheet.getRange('A' + (8 + Object.keys(byCategory).length + 1)).setFontWeight('bold');
  sumSheet.setColumnWidth(1, 250);
  sumSheet.setColumnWidth(2, 100);
}

// ============================================
// Execute Moves (Advanced)
// ============================================
function executeMoves() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recSheet = ss.getSheetByName(INVENTORY_CONFIG.SHEETS.RECOMMENDATIONS);

  if (!recSheet || recSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No recommendations to process.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Execute Moves',
    'Enter the ID of the target root folder (the company folder created by Drive Setup):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const targetRootId = response.getResponseText().trim();

  try {
    const targetRoot = DriveApp.getFolderById(targetRootId);
    const data = recSheet.getDataRange().getValues();
    let moved = 0;

    for (let i = 1; i < data.length; i++) {
      const shouldMove = data[i][5];  // Move checkbox
      if (shouldMove === true) {
        const fileId = data[i][6];
        const targetFolderName = data[i][2];

        // Find or create target subfolder
        const targetFolder = getOrCreateSubfolder(targetRoot, targetFolderName);

        // Move file
        const file = DriveApp.getFileById(fileId);
        file.moveTo(targetFolder);
        moved++;

        Logger.log('Moved: ' + data[i][0] + ' to ' + targetFolderName);
      }
    }

    ui.alert('✅ Moved ' + moved + ' files successfully!');

  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

function getOrCreateSubfolder(parent, folderName) {
  // Handle nested paths like "Listings/Photos & Media"
  const parts = folderName.split('/');
  let current = parent;

  for (const part of parts) {
    const folders = current.getFoldersByName(part);
    if (folders.hasNext()) {
      current = folders.next();
    } else {
      current = current.createFolder(part);
    }
  }

  return current;
}
