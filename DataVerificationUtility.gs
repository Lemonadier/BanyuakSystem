/**
 * ClassBank Data Verification & Migration Utility
 * 
 * This script helps you:
 * 1. Verify all required sheets exist with correct headers
 * 2. Check for any misplaced data (e.g., attendance data in Transactions sheet)
 * 3. Optionally migrate mixed data to the correct sheets
 * 
 * HOW TO USE:
 * 1. Open your Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code in a new file (e.g., "DataUtility.gs")
 * 4. Run the desired function from the toolbar
 */

const SHEET_ID = "1I3H67ou-hG1kKpbJfcNQRKA17j8CLGMm6tInVUd-2S8";

/**
 * Main verification function - Run this first to check your setup
 */
function verifySheetStructure() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const log = [];
  
  log.push("=== ClassBank Sheet Structure Verification ===\n");
  
  // Check all required sheets
  const requiredSheets = {
    "Students": ["Student ID", "Name", "Grade", "No", "Created At"],
    "Transactions": ["Transaction ID", "Student ID", "Type", "Amount", "Date", "Timestamp", "Note"],
    "Attendance": ["Transaction ID", "Student ID", "Status", "Date", "Timestamp"],
    "Health": ["Transaction ID", "Student ID", "Weight", "Height", "BMI", "Date", "Timestamp"],
    "Profile": ["Transaction ID", "Student ID", "Mood", "Score", "Date", "Timestamp"],
    "Settings": ["Key", "Value", "Updated At"]
  };
  
  Object.keys(requiredSheets).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      log.push(`❌ MISSING: Sheet "${sheetName}" does not exist`);
    } else {
      const rowCount = sheet.getLastRow();
      if (rowCount === 0) {
        log.push(`⚠️ EMPTY: Sheet "${sheetName}" exists but has no data`);
      } else {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const expectedHeaders = requiredSheets[sheetName];
        const headersMatch = JSON.stringify(headers.slice(0, expectedHeaders.length)) === JSON.stringify(expectedHeaders);
        
        if (headersMatch) {
          log.push(`✅ OK: Sheet "${sheetName}" - ${rowCount - 1} rows of data`);
        } else {
          log.push(`⚠️ HEADERS: Sheet "${sheetName}" has incorrect headers`);
          log.push(`   Expected: [${expectedHeaders.join(", ")}]`);
          log.push(`   Found: [${headers.join(", ")}]`);
        }
      }
    }
  });
  
  // Check for potential data mixing in Transactions sheet
  log.push("\n=== Checking for Mixed Data Types ===\n");
  const txSheet = ss.getSheetByName("Transactions");
  if (txSheet && txSheet.getLastRow() > 1) {
    const data = txSheet.getRange(2, 1, txSheet.getLastRow() - 1, txSheet.getLastColumn()).getValues();
    let depositCount = 0, withdrawCount = 0, otherCount = 0;
    
    data.forEach(row => {
      const type = row[2]; // Type column
      if (type === 'Deposit') depositCount++;
      else if (type === 'Withdraw') withdrawCount++;
      else if (type) otherCount++;
    });
    
    log.push(`Transactions Sheet Analysis:`);
    log.push(`  - Deposits: ${depositCount}`);
    log.push(`  - Withdrawals: ${withdrawCount}`);
    if (otherCount > 0) {
      log.push(`  - ⚠️ OTHER TYPES: ${otherCount} (may be misplaced attendance/health/profile data)`);
    } else {
      log.push(`  - ✅ No mixed data types detected`);
    }
  }
  
  // Display results
  const result = log.join("\n");
  Logger.log(result);
  
  // Also show in UI
  SpreadsheetApp.getUi().alert("Sheet Structure Verification Complete", result, SpreadsheetApp.getUi().ButtonSet.OK);
  
  return result;
}

/**
 * Count records by type across all sheets
 */
function generateDataReport() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const report = [];
  
  report.push("=== ClassBank Data Report ===\n");
  
  // Students
  const studentsSheet = ss.getSheetByName("Students");
  const studentCount = studentsSheet ? studentsSheet.getLastRow() - 1 : 0;
  report.push(`👨‍🎓 Students: ${Math.max(0, studentCount)}`);
  
  // Bank Transactions
  const txSheet = ss.getSheetByName("Transactions");
  const txCount = txSheet ? txSheet.getLastRow() - 1 : 0;
  report.push(`💰 Bank Transactions: ${Math.max(0, txCount)}`);
  
  // Attendance
  const attSheet = ss.getSheetByName("Attendance");
  const attCount = attSheet ? attSheet.getLastRow() - 1 : 0;
  report.push(`📅 Attendance Records: ${Math.max(0, attCount)}`);
  
  // Health
  const healthSheet = ss.getSheetByName("Health");
  const healthCount = healthSheet ? healthSheet.getLastRow() - 1 : 0;
  report.push(`❤️ Health Measurements: ${Math.max(0, healthCount)}`);
  
  // Profile
  const profileSheet = ss.getSheetByName("Profile");
  const profileCount = profileSheet ? profileSheet.getLastRow() - 1 : 0;
  report.push(`📊 Profile Entries: ${Math.max(0, profileCount)}`);
  
  const result = report.join("\n");
  Logger.log(result);
  SpreadsheetApp.getUi().alert("Data Report", result, SpreadsheetApp.getUi().ButtonSet.OK);
  
  return result;
}

/**
 * Create missing sheets with proper headers
 */
function createMissingSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const created = [];
  
  const sheetTemplates = {
    "Students": {
      headers: ["Student ID", "Name", "Grade", "No", "Created At"],
      color: "#f3f4f6"
    },
    "Transactions": {
      headers: ["Transaction ID", "Student ID", "Type", "Amount", "Date", "Timestamp", "Note"],
      color: "#e0e7ff"
    },
    "Attendance": {
      headers: ["Transaction ID", "Student ID", "Status", "Date", "Timestamp"],
      color: "#d1fae5"
    },
    "Health": {
      headers: ["Transaction ID", "Student ID", "Weight", "Height", "BMI", "Date", "Timestamp"],
      color: "#fce7f3"
    },
    "Profile": {
      headers: ["Transaction ID", "Student ID", "Mood", "Score", "Date", "Timestamp"],
      color: "#fef3c7"
    },
    "Settings": {
      headers: ["Key", "Value", "Updated At"],
      color: "#e2e8f0"
    }
  };
  
  Object.keys(sheetTemplates).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const template = sheetTemplates[sheetName];
      sheet.appendRow(template.headers);
      sheet.getRange(1, 1, 1, template.headers.length)
        .setFontWeight("bold")
        .setBackground(template.color);
      created.push(sheetName);
    }
  });
  
  if (created.length > 0) {
    const msg = `Created missing sheets:\n- ${created.join("\n- ")}`;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert("Sheets Created", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("All Sheets Exist", "All required sheets already exist!", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Add a custom menu to the spreadsheet for easy access
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ClassBank Utilities')
    .addItem('1️⃣ Verify Sheet Structure', 'verifySheetStructure')
    .addItem('2️⃣ Generate Data Report', 'generateDataReport')
    .addItem('3️⃣ Create Missing Sheets', 'createMissingSheets')
    .addToUi();
}
