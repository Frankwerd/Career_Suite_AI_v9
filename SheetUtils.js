// File: SheetUtils.gs
// Project: CareerSuite.AI Job Tracker
// Description: Contains utility functions for Google Sheets interaction,
// including sheet creation, formatting, and data access setup.

/**
 * Converts a 1-based column index to its letter representation (e.g., 1 -> A, 27 -> AA).
 * @param {number} column The 1-based column index.
 * @return {string} The column letter(s).
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Sets up basic formatting for a given sheet: headers, frozen row, column widths, and optional banding.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to format.
 * @param {Array<string>} headersArray An array of strings for the header row.
 * @param {Array<{col: number, width: number}>} [columnWidthsArray] Optional. Array of objects specifying column index (1-based) and width.
 * @param {boolean} [applyBandingFlag] Optional. True to apply row banding. Defaults to false.
 * @param {string} [bandingHeaderColorHex] Optional. Hex color for banding header. Defaults to BRAND_COLORS.CAROLINA_BLUE.
 * @param {string} [bandingFirstRowColorHex] Optional. Hex color for first band. Defaults to BRAND_COLORS.WHITE.
 * @param {string} [bandingSecondRowColorHex] Optional. Hex color for second band. Defaults to BRAND_COLORS.PALE_GREY.
 * @return {boolean} True if formatting was successful, false otherwise.
 */
function setupSheetFormatting(sheet, headersArray, columnWidthsArray, applyBandingFlag, bandingThemeEnum) {
  const FUNC_NAME = "setupSheetFormatting";

  if (!sheet || typeof sheet.getName !== 'function') {
    Logger.log(`[${FUNC_NAME} ERROR] Invalid sheet object passed. Sheet: ${sheet}`);
    return false;
  }
  
  let effectiveHeaderCount = 0;
  if (headersArray && Array.isArray(headersArray) && headersArray.length > 0) {
      effectiveHeaderCount = headersArray.length;
  } else if (sheet.getLastColumn() > 0) {
      effectiveHeaderCount = sheet.getLastColumn();
      Logger.log(`[${FUNC_NAME} WARN] No headersArray for "${sheet.getName()}". Using lastCol (${effectiveHeaderCount}) for effective header count.`);
  } else {
      Logger.log(`[${FUNC_NAME} WARN] Sheet "${sheet.getName()}" empty and no headers. Formatting limited.`);
      // If headersArray is strictly required for other steps, you might return false here.
      // For now, allow it to try setting frozen rows at least.
  }

  Logger.log(`[${FUNC_NAME} INFO] Applying formatting to sheet: "${sheet.getName()}". Effective headers count: ${effectiveHeaderCount}`);
  
  try {
    // --- 0. Clear existing formats on a substantial range to prevent conflicts ---
    // This needs to happen BEFORE setting new headers if they might have old formats.
    const rowsToClearFormat = Math.min(sheet.getMaxRows(), Math.max(200, sheet.getLastRow() > 1 ? sheet.getLastRow() : 200) + 50);
    if (rowsToClearFormat > 0 && effectiveHeaderCount > 0) {
      sheet.getRange(1, 1, rowsToClearFormat, effectiveHeaderCount).clearFormat().removeCheckboxes();
      Logger.log(`[${FUNC_NAME} INFO] Cleared formats on range 1:${rowsToClearFormat}, cols 1:${effectiveHeaderCount} on "${sheet.getName()}".`);
    } else if (rowsToClearFormat === 0 && effectiveHeaderCount === 0 && sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearFormat().removeCheckboxes();
      Logger.log(`[${FUNC_NAME} INFO] Sheet "${sheet.getName()}" empty, cleared all formats.`);
    }

    // --- 1. Set Headers (only if headersArray is valid) ---
    if (headersArray && Array.isArray(headersArray) && headersArray.length > 0) {
      const headerRange = sheet.getRange(1, 1, 1, headersArray.length);
      headerRange.setValues([headersArray]);
      headerRange.setFontWeight('bold')
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle')
                 .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setRowHeight(1, 45); 
      Logger.log(`[${FUNC_NAME} INFO] Headers set for "${sheet.getName()}".`);
    } else {
        Logger.log(`[${FUNC_NAME} INFO] No headersArray provided, skipping explicit header setting for "${sheet.getName()}".`);
    }

    // --- 2. Freeze Top Row ---
    if (sheet.getFrozenRows() !== 1) {
      try { sheet.setFrozenRows(1); } catch(e) { Logger.log(`[${FUNC_NAME} WARN] Could not set frozen rows on "${sheet.getName()}": ${e.message}`);}
    }

    // --- 3. Set Column Widths ---
    if (columnWidthsArray && Array.isArray(columnWidthsArray) && effectiveHeaderCount > 0) {
      columnWidthsArray.forEach(cw => {
        try {
          if (cw.col > 0 && cw.width > 0 && cw.col <= effectiveHeaderCount) {
            sheet.setColumnWidth(cw.col, cw.width);
          }
        } catch (e) { Logger.log(`[${FUNC_NAME} WARN] Error setting width for col ${cw.col} on "${sheet.getName()}": ${e.message}`); }
      });
      Logger.log(`[${FUNC_NAME} INFO] Column widths applied to "${sheet.getName()}".`);
    }

    // --- 4. General Data Area Formatting ---
    if (effectiveHeaderCount > 0) {
        const lastDataRowForFormat = sheet.getLastRow();
        const numRowsToFormatDataArea = Math.min(sheet.getMaxRows() -1, Math.max(199, lastDataRowForFormat > 1 ? lastDataRowForFormat -1 : 199)); // Format data rows starting from row 2
        if (numRowsToFormatDataArea > 0) { // Check if there's at least one data row to format
            const dataRange = sheet.getRange(2, 1, numRowsToFormatDataArea, effectiveHeaderCount);
            try {
                dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top');
            } catch(e) {Logger.log(`[${FUNC_NAME} WARN] Error applying data area wrap/align on "${sheet.getName()}": ${e.message}`);}
        }
    }
    
    // --- 5. Apply Banding (if requested) ---
    if (applyBandingFlag && effectiveHeaderCount > 0) {
      const lastPopulatedRow = sheet.getLastRow(); // Get current last row with any content
      const defaultVisualRowCountForBanding = 200; // How many rows to show banded if sheet is new/empty

      // Determine how many rows to include in the banding range
      // If sheet has more content than default, use that. Otherwise, use default.
      // Always ensure it's at least 2 (header + 1 data row) for applyRowBanding to have context.
      let bandingTotalRows;
      if (lastPopulatedRow <= 1) { // Sheet is empty or only has header
        bandingTotalRows = defaultVisualRowCountForBanding;
      } else {
        bandingTotalRows = Math.max(lastPopulatedRow, defaultVisualRowCountForBanding);
      }
      bandingTotalRows = Math.max(2, bandingTotalRows); // Ensure at least 2 rows
      bandingTotalRows = Math.min(bandingTotalRows, sheet.getMaxRows()); // Don't exceed max sheet rows

      const rangeForBanding = sheet.getRange(1, 1, bandingTotalRows, effectiveHeaderCount);
      
      try {
        // It's already been cleared comprehensively above, so direct apply.
        // No, still clear bandings directly on THE sheet object before applying to a range
        const existingBandingsOnSheet = sheet.getBandings();
        for (let i = 0; i < existingBandingsOnSheet.length; i++) {
            existingBandingsOnSheet[i].remove();
        }
        Logger.log(`[${FUNC_NAME} INFO] Cleared ALL existing bandings on sheet "${sheet.getName()}". Attempting new banding on ${rangeForBanding.getA1Notation()}`);
        
        const themeToApply = bandingThemeEnum || SpreadsheetApp.BandingTheme.LIGHT_GREY; 
        const banding = rangeForBanding.applyRowBanding(themeToApply, true, false); // Header row true, footer false
        
        Logger.log(`[${FUNC_NAME} INFO] Banding with theme "${themeToApply.toString()}" applied to sheet "${sheet.getName()}", range: ${rangeForBanding.getA1Notation()}.`);
      } catch (eBanding) {
        Logger.log(`[${FUNC_NAME} WARN] BANDING ATTEMPT FAILED on sheet "${sheet.getName()}": ${eBanding.toString()}. Range: ${rangeForBanding.getA1Notation()}. Theme: ${bandingThemeEnum ? bandingThemeEnum.toString() : 'Default'}`);
      }
    }

    // --- 6. Delete Extra Columns ---
    if (effectiveHeaderCount > 0) {
        const maxSheetCols = sheet.getMaxColumns();
        if (maxSheetCols > effectiveHeaderCount) {
            try { 
                sheet.deleteColumns(effectiveHeaderCount + 1, maxSheetCols - effectiveHeaderCount); 
                Logger.log(`[${FUNC_NAME} INFO] Extra columns ${effectiveHeaderCount + 1}-${maxSheetCols} removed from "${sheet.getName()}".`);
            } catch(e){ Logger.log(`[${FUNC_NAME} WARN] Error removing unused columns on "${sheet.getName()}": ${e.message}`);  }
        }
    }
    return true;
  } catch (err) {
    Logger.log(`[${FUNC_NAME} ERROR] Major error during formatting of sheet "${sheet.getName()}": ${err.toString()}\nStack: ${err.stack}`);
    return false;
  }
}

// The getOrCreateSpreadsheetAndSheet function remains the same as the corrected version I provided previously
// (the one NOT named _Fallback)
function getOrCreateSpreadsheetAndSheet() {
  let ss = null;
  const FUNC_NAME = "getOrCreateSpreadsheetAndSheet"; 

  if (FIXED_SPREADSHEET_ID && FIXED_SPREADSHEET_ID.trim() !== "" && FIXED_SPREADSHEET_ID !== "YOUR_MASTER_TEMPLATE_SHEET_ID_GOES_HERE") {
    Logger.log(`[${FUNC_NAME} INFO] Attempting to open by Fixed ID: "${FIXED_SPREADSHEET_ID}"`);
    try { ss = SpreadsheetApp.openById(FIXED_SPREADSHEET_ID); Logger.log(`[${FUNC_NAME} INFO] Opened "${ss.getName()}" via Fixed ID.`); }
    catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Fixed ID Fail: ${e.message}. Trying by name.`); ss = null; }
  }
  
  if (!ss) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getName() === TARGET_SPREADSHEET_FILENAME) { 
        ss = active; Logger.log(`[${FUNC_NAME} INFO] Using Active Spreadsheet: "${ss.getName()}".`);
    } else {
        Logger.log(`[${FUNC_NAME} INFO] Finding/creating by name: "${TARGET_SPREADSHEET_FILENAME}".`);
        try {
          const files = DriveApp.getFilesByName(TARGET_SPREADSHEET_FILENAME);
          if (files.hasNext()) {
            ss = SpreadsheetApp.open(files.next()); Logger.log(`[${FUNC_NAME} INFO] Found existing: "${ss.getName()}".`);
          } else {
            ss = SpreadsheetApp.create(TARGET_SPREADSHEET_FILENAME); Logger.log(`[${FUNC_NAME} INFO] Created new: "${ss.getName()}".`);
          }
        } catch (eDrive) { Logger.log(`[${FUNC_NAME} FATAL] Drive/Open/Create Fail: ${eDrive.message}.`); return { spreadsheet: null }; }
    }
  }
  if (!ss) { Logger.log(`[${FUNC_NAME} FATAL] Spreadsheet object is null.`); }
  return { spreadsheet: ss };
}

/**
 * Gets or creates the main project spreadsheet. Fallback for manual runs.
 * @return {{spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null}}
 */
function getOrCreateSpreadsheetAndSheet() {
  let ss = null;
  const FUNC_NAME = "getOrCreateSpreadsheetAndSheet";

  if (FIXED_SPREADSHEET_ID && FIXED_SPREADSHEET_ID.trim() !== "" && FIXED_SPREADSHEET_ID !== "YOUR_MASTER_TEMPLATE_SHEET_ID_GOES_HERE") {
    Logger.log(`[${FUNC_NAME} INFO] Attempting to open by Fixed ID: "${FIXED_SPREADSHEET_ID}"`);
    try { ss = SpreadsheetApp.openById(FIXED_SPREADSHEET_ID); Logger.log(`[${FUNC_NAME} INFO] Opened "${ss.getName()}" via Fixed ID.`); }
    catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Fixed ID Fail: ${e.message}. Trying by name.`); ss = null; }
  }
  
  if (!ss) {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active && active.getName() === TARGET_SPREADSHEET_FILENAME) { 
        ss = active; Logger.log(`[${FUNC_NAME} INFO] Using Active Spreadsheet: "${ss.getName()}".`);
    } else {
        Logger.log(`[${FUNC_NAME} INFO] Finding/creating by name: "${TARGET_SPREADSHEET_FILENAME}".`);
        try {
          const files = DriveApp.getFilesByName(TARGET_SPREADSHEET_FILENAME);
          if (files.hasNext()) {
            ss = SpreadsheetApp.open(files.next()); Logger.log(`[${FUNC_NAME} INFO] Found existing: "${ss.getName()}".`);
          } else {
            ss = SpreadsheetApp.create(TARGET_SPREADSHEET_FILENAME); Logger.log(`[${FUNC_NAME} INFO] Created new: "${ss.getName()}".`);
          }
        } catch (eDrive) { Logger.log(`[${FUNC_NAME} FATAL] Drive/Open/Create Fail: ${eDrive.message}.`); return { spreadsheet: null }; }
    }
  }
  if (!ss) { Logger.log(`[${FUNC_NAME} FATAL] Spreadsheet object is null.`); }
  return { spreadsheet: ss };
}
