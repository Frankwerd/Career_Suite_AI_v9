// File: Main.gs
// Project: CareerSuite.AI Job Tracker
// Description: Main script file orchestrating setup, email processing, and UI for the job application tracker.
// Author: Francis John LiButti (Originals), AI Integration & Refinements by Assistant
// Version: 7

/**
 * Runs the complete initial setup for ALL modules of the CareerSuite.AI Job Tracker.
 * This function is designed to be called by the WebApp after a new sheet is created for the user,
 * or can be run manually from the Apps Script editor or a custom menu.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [passedSpreadsheet] Optional. The spreadsheet object to set up.
 *        If not provided (e.g., manual run), it attempts to get/create the sheet based on Config.gs.
 * @return {{success: boolean, message: string, detailedMessages: Array<string>, sheetId: string | null, sheetUrl: string | null}}
 *         An object indicating the outcome of the setup process.
 */
function runFullProjectInitialSetup(passedSpreadsheet) {
  const RUNDATE = new Date().toISOString();
  const FUNC_NAME = "runFullProjectInitialSetup";
  Logger.log(`==== ${FUNC_NAME}: STARTING (CareerSuite.AI v1.2 - ${RUNDATE}) ====`);
  let overallSuccess = true;
  let setupMessages = [];
  let activeSS;

  if (passedSpreadsheet && typeof passedSpreadsheet.getId === 'function') {
    activeSS = passedSpreadsheet;
    Logger.log(`[${FUNC_NAME} INFO] Using PASSED spreadsheet: "${activeSS.getName()}", ID: ${activeSS.getId()}`);
  } else {
    // Fallback logic for manual runs from the editor
    Logger.log(`[${FUNC_NAME} INFO] No spreadsheet passed (e.g., manual run from editor). Attempting to get active spreadsheet.`);
    activeSS = SpreadsheetApp.getActiveSpreadsheet(); 
    // If still no activeSS (e.g. script run headless without context), then try getOrCreate
    if (!activeSS) {
        Logger.log(`[${FUNC_NAME} WARN] No specific spreadsheet passed and no active spreadsheet found. Fallback to getOrCreateSpreadsheetAndSheet(). This might target a generic or template sheet.`);
        const { spreadsheet: foundOrCreatedSS } = getOrCreateSpreadsheetAndSheet(); // From SheetUtils.gs
        activeSS = foundOrCreatedSS;
    } else {
        Logger.log(`[${FUNC_NAME} INFO] Using ACTIVE spreadsheet: "${activeSS.getName()}", ID: ${activeSS.getId()}`);
    }
  }

  if (!activeSS) {
    const errorMsg = `CRITICAL [${FUNC_NAME}]: No valid spreadsheet could be determined. Setup aborted.`;
    Logger.log(errorMsg);
    return { success: false, message: errorMsg, detailedMessages: [errorMsg], sheetId: null, sheetUrl: null };
  }

  // --- TEMPLATE CHECK ---
  // TEMPLATE_SHEET_ID must be defined in Config.js
  if (typeof TEMPLATE_SHEET_ID !== 'undefined' && TEMPLATE_SHEET_ID !== "" && activeSS.getId() === TEMPLATE_SHEET_ID) {
    const templateMsg = `[${FUNC_NAME} INFO] Target spreadsheet is the TEMPLATE (ID: ${TEMPLATE_SHEET_ID}). Setup SKIPPED to prevent modifications to the template. This is normal if setup was triggered on the template directly.`;
    Logger.log(templateMsg);
    // Optionally, show a UI alert if run manually on the template
    if (!passedSpreadsheet) { // Indicates a manual run from menu or editor
        try {
            SpreadsheetApp.getUi().alert('Template Sheet', 'Initial setup is not meant to be run on the template sheet itself. Please make a copy first, then run the setup on your new sheet if needed, or ensure the automated setup runs on your copy.', SpreadsheetApp.getUi().ButtonSet.OK);
        } catch (e) { /* UI not available, e.g. headless run, ignore */ }
    }
    return { 
        success: true, // Success because the check worked as expected
        message: "Setup skipped: Target is template.", 
        detailedMessages: [templateMsg], 
        sheetId: activeSS.getId(), 
        sheetUrl: activeSS.getUrl() 
    };
  }
  // --- END TEMPLATE CHECK ---

  // --- 1. Setup Job Application Tracker Module ---
  Logger.log(`\n[${FUNC_NAME} INFO] --- Starting Job Application Tracker Module Setup ---`);
  try {
    const trackerResult = initialSetup_LabelsAndSheet(activeSS);
    if(trackerResult.messages) setupMessages.push(...trackerResult.messages.map(m => `Tracker: ${m}`));
    if (!trackerResult.success) { overallSuccess = false; Logger.log(`[${FUNC_NAME} ERROR] Tracker Module FAILED.`);}
    else { Logger.log(`[${FUNC_NAME} INFO] Tracker Module Success.`); }
  } catch (e) {
    Logger.log(`[${FUNC_NAME} CRITICAL ERROR] Tracker Exception: ${e.toString()}\n${e.stack}`);
    setupMessages.push(`Tracker: CRITICAL EXCEPTION - ${e.message}`); overallSuccess = false;
  }

  // --- 2. Setup Job Leads Tracker Module ---
  if (typeof runInitialSetup_JobLeadsModule === "function") {
    Logger.log(`\n[${FUNC_NAME} INFO] --- Starting Job Leads Tracker Module Setup ---`);
    try {
      const leadsResult = runInitialSetup_JobLeadsModule(activeSS); // From Leads_Main.gs
      if(leadsResult.messages) setupMessages.push(...leadsResult.messages.map(m => `Leads: ${m}`));
      if (!leadsResult.success) { overallSuccess = false; Logger.log(`[${FUNC_NAME} ERROR] Leads Module FAILED.`); }
      else { Logger.log(`[${FUNC_NAME} INFO] Leads Module Success.`); }
    } catch (e) {
      Logger.log(`[${FUNC_NAME} CRITICAL ERROR] Leads Exception: ${e.toString()}\n${e.stack}`);
      setupMessages.push(`Leads: CRITICAL EXCEPTION - ${e.message}`); overallSuccess = false;
    }
  } else { Logger.log(`[${FUNC_NAME} INFO] runInitialSetup_JobLeadsModule not found. Skipping leads setup.`); }

  // --- 3. BRANDING: Final Tab Order ---
  if (overallSuccess) { 
    Logger.log(`[${FUNC_NAME} INFO] Applying final tab order...`);
    try {
        const tabOrder = [ DASHBOARD_TAB_NAME, APP_TRACKER_SHEET_TAB_NAME, LEADS_SHEET_TAB_NAME ];
        let currentPosition = 1; 
        for (const sheetName of tabOrder) {
            const sheetToMove = activeSS.getSheetByName(sheetName);
            if (sheetToMove) {
                // The actual index of the sheet (1-based for user, 0-based for API usually)
                // Sheet.getIndex() returns 1-based index.
                if (sheetToMove.getIndex() !== currentPosition) { // <<<< CORRECTED THIS LINE
                     activeSS.setActiveSheet(sheetToMove); 
                     activeSS.moveActiveSheet(currentPosition);
                     Logger.log(`[${FUNC_NAME} INFO] Moved sheet "${sheetName}" to position ${currentPosition}.`);
                } else {
                    Logger.log(`[${FUNC_NAME} INFO] Sheet "${sheetName}" already at position ${currentPosition}.`);
                }
                currentPosition++;
            } else { Logger.log(`[${FUNC_NAME} WARN] Sheet "${sheetName}" for ordering not found.`); }
            Utilities.sleep(150);
        }
        const helperSheet = activeSS.getSheetByName(HELPER_SHEET_NAME);
        if (helperSheet) {
            if (!helperSheet.isSheetHidden()) helperSheet.hideSheet();
            // To ensure it's last if other tabs were added unexpectedly
            // activeSS.setActiveSheet(helperSheet); 
            // activeSS.moveActiveSheet(activeSS.getSheets().length); 
            Logger.log(`[${FUNC_NAME} INFO] Helper sheet "${HELPER_SHEET_NAME}" hidden.`);
        }
        setupMessages.push("Branding: Tab order & helper visibility verified.");
    } catch (e) { Logger.log(`[${FUNC_NAME} WARN] Error finalizing tab order: ${e.message}\nStack: ${e.stack}`); }
  }

  const finalStatusMessage = `CareerSuite.AI Full Setup ${overallSuccess ? "completed" : "had issues"}.`;
  Logger.log(`\n==== ${FUNC_NAME} SUMMARY (SS ID: ${activeSS.getId()}) ====`);
  setupMessages.forEach(msg => Logger.log(`  - ${msg}`));
  Logger.log(`Overall Status: ${overallSuccess ? "SUCCESSFUL" : "ISSUES ENCOUNTERED"}`);
  Logger.log(`==== ${FUNC_NAME} ${overallSuccess ? "CONCLUDED" : "CONCLUDED WITH ISSUES"} ====`);

  if (!passedSpreadsheet && overallSuccess) { 
      try { SpreadsheetApp.getUi().alert('CareerSuite.AI Setup Complete', `Setup finished for "${activeSS.getName()}".\n\nDetails:\n- ${setupMessages.join('\n- ')}`.substring(0,1000), SpreadsheetApp.getUi().ButtonSet.OK); } 
      catch (e) { Logger.log(`[${FUNC_NAME} INFO] UI Alert skipped: ${e.message}`); }
  } else if (!passedSpreadsheet && !overallSuccess) {
       try { SpreadsheetApp.getUi().alert('CareerSuite.AI Setup Issues', `Setup for "${activeSS.getName()}" had issues. Check logs.\n\nSummary:\n- ${setupMessages.join('\n- ')}`.substring(0,1000), SpreadsheetApp.getUi().ButtonSet.OK); } 
       catch (e) { Logger.log(`[${FUNC_NAME} INFO] UI Alert skipped: ${e.message}`); }
  }
  return { 
    success: overallSuccess, 
    message: `Full Setup ${overallSuccess ? "completed" : "had issues"}.`, 
    detailedMessages: setupMessages, 
    sheetId: activeSS.getId(), 
    sheetUrl: activeSS.getUrl() 
  };
}


/**
 * Sets up the core Job Application Tracker: Labels, Sheets (Applications, Dashboard, Helper), and Triggers.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} activeSS The spreadsheet object to set up.
 * @return {{success: boolean, messages: Array<string>}} Result of the setup for this module.
 */
function initialSetup_LabelsAndSheet(activeSS) {
  const FUNC_NAME = "initialSetup_LabelsAndSheet";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING - Tracker Module Setup ====`);
  let messages = [];
  let moduleSuccess = true;
  let dummyDataWasAdded = false;
  let dataSh, dashboardSheet, helperSheet; // Declare here for broader scope within function

  if (!activeSS || typeof activeSS.getId !== 'function') {
    const errMsg = "CRITICAL: Invalid spreadsheet object passed."; Logger.log(`[${FUNC_NAME} ERROR] ${errMsg}`);
    return { success: false, messages: [errMsg] };
  }
  Logger.log(`[${FUNC_NAME} INFO] Operating on: "${activeSS.getName()}" (ID: ${activeSS.getId()})`);

  // --- A. Core Sheet Creation & Formatting ---
  Logger.log(`[${FUNC_NAME} INFO] Setting up core sheets for Tracker module...`);
  try {
    // A.1: "Applications" Sheet
    dataSh = activeSS.getSheetByName(APP_TRACKER_SHEET_TAB_NAME);
    if (!dataSh) {
      dataSh = activeSS.insertSheet(APP_TRACKER_SHEET_TAB_NAME);
      Logger.log(`[${FUNC_NAME} INFO] Created new sheet: "${APP_TRACKER_SHEET_TAB_NAME}".`);
    } else { 
      Logger.log(`[${FUNC_NAME} INFO] Found existing sheet: "${APP_TRACKER_SHEET_TAB_NAME}".`);
    }
    // Corrected THEME for Applications
    if (!setupSheetFormatting(dataSh, 
                              APP_TRACKER_SHEET_HEADERS,        // From Config.gs
                              APP_SHEET_COLUMN_WIDTHS,          // From Config.gs
                              true,                             // applyBandingFlag = true
                              SpreadsheetApp.BandingTheme.BLUE  // <<< ENSURE THIS IS .BLUE or .CYAN
                             )) {
        throw new Error(`Formatting failed for "${APP_TRACKER_SHEET_TAB_NAME}".`);
    }
    dataSh.setTabColor(BRAND_COLORS.LAPIS_LAZULI);
    try { // Post-formatting specific tweaks
        if (PEAK_STATUS_COL > 0 && PEAK_STATUS_COL <= dataSh.getMaxColumns() && !dataSh.isColumnHiddenByUser(PEAK_STATUS_COL)) {
            dataSh.hideColumn(dataSh.getRange(1, PEAK_STATUS_COL));
        }
        if (EMAIL_LINK_COL > 0 && dataSh.getMaxRows() > 1 && dataSh.getMaxColumns() >= EMAIL_LINK_COL) {
            dataSh.getRange(2, EMAIL_LINK_COL, dataSh.getMaxRows() - 1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        }
    } catch (ePf) { Logger.log(`[${FUNC_NAME} WARN] Post-format tweaks for Apps sheet: ${ePf.message}`);}
    messages.push(`Sheet '${APP_TRACKER_SHEET_TAB_NAME}': Setup OK. Color: Lapis Lazuli.`);

    // A.2: "Dashboard" Sheet
    dashboardSheet = getOrCreateDashboardSheet(activeSS); // From Dashboard.gs
    if (!dashboardSheet) throw new Error(`Get/Create FAILED for sheet: '${DASHBOARD_TAB_NAME}'.`);
    if (!formatDashboardSheet(dashboardSheet)) { // From Dashboard.gs
         throw new Error(`Formatting FAILED for sheet: '${DASHBOARD_TAB_NAME}'.`);
    } // Tab color for dashboard is set within getOrCreateDashboardSheet in Dashboard.gs
    messages.push(`Sheet '${DASHBOARD_TAB_NAME}': Setup OK.`);

    // A.3: "DashboardHelperData" Sheet
    helperSheet = getOrCreateHelperSheet(activeSS); // From Dashboard.gs
    if (!helperSheet) throw new Error(`Get/Create FAILED for sheet: '${HELPER_SHEET_NAME}'.`);
    // Format helper sheet (headers, no banding) using SheetUtils.gs
    if(!setupSheetFormatting(helperSheet, DASHBOARD_HELPER_HEADERS, HELPER_SHEET_COLUMN_WIDTHS, false)) { 
        throw new Error(`Basic Formatting FAILED for sheet: '${HELPER_SHEET_NAME}'.`);
    }
    // **** NEW CALL ****
    if (!setupHelperSheetFormulas(helperSheet)) { // Call from Dashboard.gs to set formulas
        throw new Error(`Setting formulas FAILED for sheet: '${HELPER_SHEET_NAME}'.`);
    }
    if (!helperSheet.isSheetHidden()) helperSheet.hideSheet();
    helperSheet.setTabColor(BRAND_COLORS.CHARCOAL); // From Config.gs
    messages.push(`Sheet '${HELPER_SHEET_NAME}': Setup OK (Headers & Formulas set). Hidden. Color: Charcoal.`);

  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] Core sheet setup failed: ${e.toString()}\nStack: ${e.stack}`);
    messages.push(`Core sheet setup FAILED: ${e.message}.`); moduleSuccess = false;
  }

  // --- B. Gmail Label & Filter Setup ---
  let trackerToProcessLabelId = null;
  if (moduleSuccess) {
    Logger.log(`[${FUNC_NAME} INFO] Setting up Gmail labels & filters for Tracker...`);
    try {
        getOrCreateLabel(MASTER_GMAIL_LABEL_PARENT); Utilities.sleep(100);       // From Config.gs
        getOrCreateLabel(TRACKER_GMAIL_LABEL_PARENT);  Utilities.sleep(100);      // From Config.gs
        const toProcessLabelObject = getOrCreateLabel(TRACKER_GMAIL_LABEL_TO_PROCESS); Utilities.sleep(100); // From Config.gs
        getOrCreateLabel(TRACKER_GMAIL_LABEL_PROCESSED); Utilities.sleep(100);   // From Config.gs
        getOrCreateLabel(TRACKER_GMAIL_LABEL_MANUAL_REVIEW); Utilities.sleep(100); // From Config.gs
        
        if (toProcessLabelObject) {
            Utilities.sleep(300);
            const advancedGmailService = Gmail; // Assumes Advanced Gmail API Service is enabled
            if (!advancedGmailService || !advancedGmailService.Users || !advancedGmailService.Users.Labels) {
                throw new Error("Advanced Gmail Service not available/enabled for label ID fetch.");
            }
            const labelsListResponse = advancedGmailService.Users.Labels.list('me');
            if (labelsListResponse.labels && labelsListResponse.labels.length > 0) {
                const targetLabelInfo = labelsListResponse.labels.find(l => l.name === TRACKER_GMAIL_LABEL_TO_PROCESS);
                if (targetLabelInfo && targetLabelInfo.id) {
                    trackerToProcessLabelId = targetLabelInfo.id;
                } else { Logger.log(`[${FUNC_NAME} WARN] Label "${TRACKER_GMAIL_LABEL_TO_PROCESS}" ID not found via Advanced Service.`); }
            } else { Logger.log(`[${FUNC_NAME} WARN} No labels returned by Advanced Gmail Service.`); }
        }
        if (!trackerToProcessLabelId) throw new Error(`CRITICAL: Could not get ID for Gmail label "${TRACKER_GMAIL_LABEL_TO_PROCESS}". Filter creation will fail.`);
        messages.push("Tracker Labels & 'To Process' ID: OK.");
        
        // Filter Creation
        const filterQuery = TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES; // from Config.gs
        const gmailApiServiceForFilter = Gmail; // Advanced Gmail Service
        let filterExists = false;
        const existingFiltersResponse = gmailApiServiceForFilter.Users.Settings.Filters.list('me');
        const existingFiltersList = (existingFiltersResponse && existingFiltersResponse.filter && Array.isArray(existingFiltersResponse.filter)) ? existingFiltersResponse.filter : [];
        
        for (const filterItem of existingFiltersList) {
            if (filterItem.criteria?.query === filterQuery && filterItem.action?.addLabelIds?.includes(trackerToProcessLabelId)) {
                filterExists = true; break;
            }
        }
        if (!filterExists) {
            const filterResource = { criteria: { query: filterQuery }, action: { addLabelIds: [trackerToProcessLabelId], removeLabelIds: ['INBOX'] } };
            const createdFilterResponse = gmailApiServiceForFilter.Users.Settings.Filters.create(filterResource, 'me');
            if (!createdFilterResponse || !createdFilterResponse.id) {
                 throw new Error(`Gmail filter creation for tracker FAILED or did not return ID. Response: ${JSON.stringify(createdFilterResponse)}`);
            }
            messages.push("Tracker Filter: CREATED.");
        } else { messages.push("Tracker Filter: Exists."); }

    } catch (e) {
        Logger.log(`[${FUNC_NAME} ERROR] Gmail Label/Filter setup: ${e.toString()}`);
        messages.push(`Gmail Label/Filter setup FAILED: ${e.message}.`); moduleSuccess = false;
    }
  }
  
  // --- C. Add Dummy Data ---
  let dummyRows = []; // To scope it for removal block
  if (moduleSuccess && dataSh && dataSh.getLastRow() <= 1) { // Only if sheet is truly empty (just header)
    Logger.log(`[${FUNC_NAME} INFO] Adding dummy data to "${APP_TRACKER_SHEET_TAB_NAME}".`);
    try {
        const today = new Date(); 
        const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000); 
        const twoWeeksAgo = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);
        dummyRows = [ // Assign to the outer scoped variable
            [new Date(), twoWeeksAgo, "LinkedIn", "DemoCorp Alpha", "Software Intern", DEFAULT_STATUS, DEFAULT_STATUS, twoWeeksAgo, "Applied to Alpha", "https://example.com/alpha", "dummyMsgIdAlpha","Initial notes for Alpha"],
            [new Date(), weekAgo, "Indeed", "Test Inc. Beta", "QA Analyst", INTERVIEW_STATUS, INTERVIEW_STATUS, weekAgo, "Interview Scheduled for Beta", "https://example.com/beta", "dummyMsgIdBeta","Follow up after Beta interview"]
            // Add a third dummy row if needed for chart variety
        ];
        dummyRows = dummyRows.map(r => {
            while(r.length < TOTAL_COLUMNS_IN_APP_SHEET) r.push(""); return r.slice(0,TOTAL_COLUMNS_IN_APP_SHEET);
        });
        dataSh.getRange(2, 1, dummyRows.length, TOTAL_COLUMNS_IN_APP_SHEET).setValues(dummyRows);
        dummyDataWasAdded = true; messages.push(`Dummy data added (${dummyRows.length} rows).`);
    } catch(e) { Logger.log(`[${FUNC_NAME} WARN] Failed adding dummy data: ${e.message}`); messages.push("Dummy data add FAILED.");}
  }

    // --- D. Update Dashboard Metrics ---
  if (moduleSuccess && dashboardSheet && helperSheet && dataSh ) { 
    Logger.log(`[${FUNC_NAME} INFO] Ensuring dashboard charts are built/updated based on formula-driven helper data...`);
    try { 
      // Pass all sheets just in case, though applicationsSheet might not be directly used by the new updateDashboardMetrics for populating helper
      updateDashboardMetrics(dashboardSheet, helperSheet, dataSh); 
      messages.push("Dashboard Charts: Update/Creation attempted."); 
    } catch (e) { 
      Logger.log(`[${FUNC_NAME} ERROR] updateDashboardMetrics call failed: ${e.toString()}`); 
      messages.push(`Dashboard Charts update FAILED: ${e.message}.`);
    }
  }

  // --- E. Remove Dummy Data ---
  if (moduleSuccess && dummyDataWasAdded && dataSh && dummyRows.length > 0) {
    Logger.log(`[${FUNC_NAME} INFO] Removing dummy data (${dummyRows.length} rows)...`);
    try { 
      if (dataSh.getLastRow() >= (1 + dummyRows.length)) { // Check if enough rows exist to delete
        dataSh.deleteRows(2, dummyRows.length); 
        messages.push("Dummy data removed."); 
      } else {
        Logger.log(`[${FUNC_NAME} WARN] Not enough rows to delete all dummy data as expected. Sheet lastRow: ${dataSh.getLastRow()}, Dummy rows: ${dummyRows.length}`);
      }
    } catch(e) { Logger.log(`[${FUNC_NAME} WARN] Failed removing dummy data: ${e.message}`); }
  }

  // --- F. Trigger Verification/Creation ---
  if(moduleSuccess) {
    Logger.log(`[${FUNC_NAME} INFO] Setting up triggers for Tracker module...`);
    try { // Assumes createTimeDrivenTrigger & createOrVerifyStaleRejectTrigger are in Triggers.gs
        if (createTimeDrivenTrigger('processJobApplicationEmails', 1)) messages.push("Trigger 'processJobApplicationEmails': CREATED."); 
        else messages.push("Trigger 'processJobApplicationEmails': Exists/Verified.");
        if (createOrVerifyStaleRejectTrigger('markStaleApplicationsAsRejected', 2)) messages.push("Trigger 'markStaleApplicationsAsRejected': CREATED."); 
        else messages.push("Trigger 'markStaleApplicationsAsRejected': Exists/Verified.");
    } catch(e) {
        Logger.log(`[${FUNC_NAME} ERROR] Trigger setup failed: ${e.toString()}`);
        messages.push(`Trigger setup FAILED: ${e.message}.`);
        moduleSuccess = false; 
    }
  } else {
    messages.push("Triggers for Tracker Module: SKIPPED due to earlier failures.");
  }

  Logger.log(`\n==== ${FUNC_NAME} ${moduleSuccess ? "COMPLETED." : "ISSUES."} ====`);
  return { success: moduleSuccess, messages: messages };
}

// --- Main Email Processing Function (Job Application Tracker) ---
function processJobApplicationEmails() {
  const FUNC_NAME = "processJobApplicationEmails";
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== ${FUNC_NAME}: STARTING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  // --- 1. Configuration & Get Spreadsheet/Sheet ---
  const userProperties = PropertiesService.getUserProperties(); // <<< CORRECTED
  const geminiApiKey = userProperties.getProperty(GEMINI_API_KEY_PROPERTY); // GEMINI_API_KEY_PROPERTY from Config.gs
  let useGemini = false;

  // Add detailed logging for the key retrieval
  if (geminiApiKey) {
    Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Retrieved key for "${GEMINI_API_KEY_PROPERTY}" from UserProperties. Value (masked): ${geminiApiKey.substring(0,4)}...${geminiApiKey.substring(geminiApiKey.length-4)}`);
  } else {
    Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] NO key found for "${GEMINI_API_KEY_PROPERTY}" in UserProperties.`);
  }

  if (geminiApiKey && geminiApiKey.trim() !== "" && geminiApiKey.startsWith("AIza") && geminiApiKey.length > 30) {
    useGemini = true;
    Logger.log(`[${FUNC_NAME} INFO] Gemini API Key found in UserProperties and appears valid. AI parsing enabled.`);
  } else {
    Logger.log(`[${FUNC_NAME} WARN] Gemini API Key from UserProperties missing or invalid. Fallback to regex parsing.`);
    if (!geminiApiKey) {
        Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Reason: geminiApiKey is null or undefined after fetching from UserProperties.`);
    } else {
        Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Reason: Key found in UserProperties but failed validation. Length: ${geminiApiKey.length}, StartsWith AIza: ${geminiApiKey.startsWith("AIza")}`);
    }
  }

  const { spreadsheet: ss } = getOrCreateSpreadsheetAndSheet(); // From SheetUtils.gs
  if (!ss) {
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Main spreadsheet could not be accessed. Aborting.`);
    return;
  }

  let dataSheet; // This is the "Applications" sheet
  try {
    dataSheet = ss.getSheetByName(APP_TRACKER_SHEET_TAB_NAME); // From Config.gs
    if (!dataSheet) {
        Logger.log(`[${FUNC_NAME} WARN] Core data tab "${APP_TRACKER_SHEET_TAB_NAME}" not found in "${ss.getName()}". Attempting to create and format it now...`);
        dataSheet = ss.insertSheet(APP_TRACKER_SHEET_TAB_NAME);
        if (!setupSheetFormatting(dataSheet, APP_TRACKER_SHEET_HEADERS, APP_SHEET_COLUMN_WIDTHS, true, SpreadsheetApp.BandingTheme.BLUE)) {
            Logger.log(`[${FUNC_NAME} FATAL ERROR] Failed to create and format the missing "${APP_TRACKER_SHEET_TAB_NAME}" tab during processing. Aborting.`);
            return;
        }
        dataSheet.setTabColor(BRAND_COLORS.LAPIS_LAZULI); // From Config.gs
        Logger.log(`[${FUNC_NAME} INFO] Created and formatted missing tab: "${APP_TRACKER_SHEET_TAB_NAME}".`);
    }
  } catch (e) {
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Error accessing/creating core data tab "${APP_TRACKER_SHEET_TAB_NAME}": ${e.message}. Aborting.`);
    return;
  }
  Logger.log(`[${FUNC_NAME} INFO] Using Spreadsheet: "${ss.getName()}", Data Tab: "${dataSheet.getName()}"`);

  // --- Get Gmail Labels ---
  let procLbl, processedLblObj, manualLblObj;
  try {
    procLbl = GmailApp.getUserLabelByName(TRACKER_GMAIL_LABEL_TO_PROCESS);      
    processedLblObj = GmailApp.getUserLabelByName(TRACKER_GMAIL_LABEL_PROCESSED); 
    manualLblObj = GmailApp.getUserLabelByName(TRACKER_GMAIL_LABEL_MANUAL_REVIEW);
    if (!procLbl) throw new Error(`Processing label "${TRACKER_GMAIL_LABEL_TO_PROCESS}" not found.`);
    if (!processedLblObj) throw new Error(`Processed label "${TRACKER_GMAIL_LABEL_PROCESSED}" not found.`);
    if (!manualLblObj) throw new Error(`Manual review label "${TRACKER_GMAIL_LABEL_MANUAL_REVIEW}" not found.`);
    Logger.log(`[${FUNC_NAME} INFO] Core Gmail labels for tracker verified.`);
  } catch(e) {
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Tracker labels missing or error fetching: ${e.message}. Aborting.`); 
    return;
  }

  // --- Preload Existing Data from "Applications" Sheet ---
  const lastR = dataSheet.getLastRow(); 
  const existingDataCache = {}; 
  const processedEmailIds = new Set();
  if (lastR >= 2) { 
    Logger.log(`[${FUNC_NAME} INFO] Preloading existing data from "${dataSheet.getName()}" (Rows 2 to ${lastR})...`);
    try {
      const colsToPreloadIndices = [COMPANY_COL, JOB_TITLE_COL, EMAIL_ID_COL, STATUS_COL, PEAK_STATUS_COL]; // From Config.gs
      const minColToRead = Math.min(...colsToPreloadIndices); const maxColToRead = Math.max(...colsToPreloadIndices);
      const numColsToRead = maxColToRead - minColToRead + 1;
      if (numColsToRead < 1 || minColToRead < 1) throw new Error("Invalid preload column calculation.");

      const preloadValues = dataSheet.getRange(2, minColToRead, lastR - 1, numColsToRead).getValues();
      const coIdx=COMPANY_COL-minColToRead, tiIdx=JOB_TITLE_COL-minColToRead, idIdx=EMAIL_ID_COL-minColToRead, stIdx=STATUS_COL-minColToRead, pkIdx=PEAK_STATUS_COL-minColToRead;

      for (let i=0; i<preloadValues.length; i++) {
        const rN=i+2, rD=preloadValues[i];
        const eId=rD[idIdx]?String(rD[idIdx]).trim():"", oCo=rD[coIdx]?String(rD[coIdx]).trim():"", oTi=rD[tiIdx]?String(rD[tiIdx]).trim():"", cS=rD[stIdx]?String(rD[stIdx]).trim():"", cPkS=rD[pkIdx]?String(rD[pkIdx]).trim():"";
        if(eId) processedEmailIds.add(eId);
        const cL=oCo.toLowerCase(); if(cL && cL!==MANUAL_REVIEW_NEEDED.toLowerCase() && cL!=='n/a'){ if(!existingDataCache[cL])existingDataCache[cL]=[]; existingDataCache[cL].push({row:rN,emailId:eId,company:oCo,title:oTi,status:cS, peakStatus:cPkS});}
      }
      Logger.log(`[${FUNC_NAME} INFO] Preload complete. Cached ${Object.keys(existingDataCache).length} companies, ${processedEmailIds.size} processed email IDs.`);
    } catch (e) { Logger.log(`[${FUNC_NAME} FATAL ERROR] Preloading data: ${e.toString()}\nStack:${e.stack}. Aborting.`); return; }
  } else { Logger.log(`[${FUNC_NAME} INFO] Applications sheet empty. No data preloaded.`); }

  // --- Fetch and Filter Emails ---
  const THREAD_PROCESSING_LIMIT = 20; 
  let threadsToProcess = [];
  try { 
    threadsToProcess = procLbl.getThreads(0, THREAD_PROCESSING_LIMIT); 
    Logger.log(`[${FUNC_NAME} DEBUG_EMAIL_FETCH] Fetched ${threadsToProcess.length} threads from label "${procLbl.getName()}".`);
  } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Failed gather threads: ${e.message}`); return; }

  const messagesToSort = []; let skippedKnownProcessedCount = 0; let messageFetchErrorCount = 0;
  for (const thread of threadsToProcess) {
    const threadId = thread.getId(); try {
      const messagesInThread = thread.getMessages();
      for (const msg of messagesInThread) {
        const msgId = msg.getId(); if (!processedEmailIds.has(msgId)) messagesToSort.push({ message: msg, date: msg.getDate(), threadId: threadId }); else skippedKnownProcessedCount++; 
      }
    } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Gather messages from thread ${threadId}: ${e.message}`); messageFetchErrorCount++; }
  }
  Logger.log(`[${FUNC_NAME} DEBUG_EMAIL_FETCH] Messages to sort = ${messagesToSort.length}, Skipped = ${skippedKnownProcessedCount}, Fetch errors = ${messageFetchErrorCount}.`);

  if (messagesToSort.length === 0) {
    Logger.log(`[${FUNC_NAME} INFO] No new unread/unprocessed messages found in label "${procLbl.getName()}".`);
    try { if (typeof updateDashboardMetrics === "function") updateDashboardMetrics(ss.getSheetByName(DASHBOARD_TAB_NAME), ss.getSheetByName(HELPER_SHEET_NAME), dataSheet); } catch (e_dash) { Logger.log(`[${FUNC_NAME} WARN] Dashboard update (no new msgs) failed: ${e_dash.message}`); }
    Logger.log(`==== ${FUNC_NAME} FINISHED (${new Date().toLocaleString()}) - No new messages. ====`);
    return;
  }
  messagesToSort.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
  Logger.log(`[${FUNC_NAME} INFO] Sorted ${messagesToSort.length} new messages.`);
  
  // --- Process Each Message (FULL LOGIC REINSTATED) ---
  let threadProcessingOutcomes = {}; 
  let processedThisRunCount = 0; 
  let sheetUpdateSuccessCount = 0; 
  let newEntryCount = 0; 
  let processingErrorCount = 0;

  for (let i = 0; i < messagesToSort.length; i++) {
    const elapsedTime = (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000;
    if (elapsedTime > 320) { 
      Logger.log(`[${FUNC_NAME} WARN] Execution time limit nearing (${elapsedTime}s). Stopping message processing loop.`); 
      break; 
    }

    const entry = messagesToSort[i];
    const { message, date: emailDateObj, threadId } = entry;
    const emailDate = new Date(emailDateObj); 
    const msgId = message.getId();
    const processingStartTimeMsg = new Date();
    if(DEBUG_MODE) Logger.log(`\n--- [${FUNC_NAME}] Processing Msg ${i+1}/${messagesToSort.length} (ID: ${msgId}, Thread: ${threadId}) ---`);

    let companyName = MANUAL_REVIEW_NEEDED, jobTitle = MANUAL_REVIEW_NEEDED, applicationStatus = null; 
    let plainBodyText = null, requiresManualReview = false, sheetWriteOpSuccessThisMessage = false;

    try {
      const emailSubject = message.getSubject() || "";
      const senderEmail = message.getFrom() || "";
      const emailPermaLink = `https://mail.google.com/mail/u/0/#inbox/${msgId}`;
      const currentTimestamp = new Date();
      let detectedPlatform = DEFAULT_PLATFORM; 
      try {
        const emailAddressMatch = senderEmail.match(/<([^>]+)>/);
        if (emailAddressMatch && emailAddressMatch[1]) {
          const senderDomain = emailAddressMatch[1].split('@')[1]?.toLowerCase();
          if (senderDomain) {
            for (const keyword in PLATFORM_DOMAIN_KEYWORDS) {
              if (senderDomain.includes(keyword)) { detectedPlatform = PLATFORM_DOMAIN_KEYWORDS[keyword]; break; }
            }
          }
        }
        if(DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Detected Platform: "${detectedPlatform}"`);
      } catch (ePlat) { Logger.log(`[${FUNC_NAME} WARN] Platform detection error: ${ePlat.message}`); }
      
      try { plainBodyText = message.getPlainBody(); } 
      catch (eBody) { Logger.log(`[${FUNC_NAME} WARN] Get Plain Body Failed for Msg ${msgId}: ${eBody.message}`); plainBodyText = ""; }

      if (useGemini && plainBodyText && plainBodyText.trim() !== "") {
        const geminiResult = callGemini_forApplicationDetails(emailSubject, plainBodyText, geminiApiKey); 
        if (geminiResult) { 
            companyName = geminiResult.company || MANUAL_REVIEW_NEEDED; 
            jobTitle = geminiResult.title || MANUAL_REVIEW_NEEDED; 
            applicationStatus = geminiResult.status;
            Logger.log(`[${FUNC_NAME} INFO] Gemini: C:"${companyName}", T:"${jobTitle}", S:"${applicationStatus}"`);
            if (!applicationStatus || applicationStatus === MANUAL_REVIEW_NEEDED || applicationStatus === "Update/Other") {
                const keywordStatus = parseBodyForStatus(plainBodyText); 
                if (keywordStatus && keywordStatus !== DEFAULT_STATUS) applicationStatus = keywordStatus;
                else if (!applicationStatus && keywordStatus === DEFAULT_STATUS) applicationStatus = DEFAULT_STATUS;
            }
        } else { 
            Logger.log(`[${FUNC_NAME} WARN] Gemini call failed for Msg ${msgId}. Fallback regex.`);
            const regexResult = extractCompanyAndTitle(message, detectedPlatform, emailSubject, plainBodyText); 
            companyName = regexResult.company; jobTitle = regexResult.title;
            applicationStatus = parseBodyForStatus(plainBodyText);
        }
      } else { 
          const regexResult = extractCompanyAndTitle(message, detectedPlatform, emailSubject, plainBodyText);
          companyName = regexResult.company; jobTitle = regexResult.title;
          applicationStatus = parseBodyForStatus(plainBodyText);
          if(DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Regex Parse: C:"${companyName}", T:"${jobTitle}", S:"${applicationStatus}"`);
      }
      
      requiresManualReview = (companyName === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED);
      const finalStatusToSet = applicationStatus || DEFAULT_STATUS;
      const companyCacheKey = (companyName !== MANUAL_REVIEW_NEEDED) ? companyName.toLowerCase() : `_manual_review_placeholder_${msgId}`;
      let existingRowInfoToUpdate = null; let targetSheetRowForUpdate = -1;

      if (companyName !== MANUAL_REVIEW_NEEDED && existingDataCache[companyCacheKey]) {
          const potentialMatches = existingDataCache[companyCacheKey];
          if (jobTitle !== MANUAL_REVIEW_NEEDED) existingRowInfoToUpdate = potentialMatches.find(e => e.title && e.title.toLowerCase() === jobTitle.toLowerCase());
          if (!existingRowInfoToUpdate && potentialMatches.length > 0) existingRowInfoToUpdate = potentialMatches.reduce((latest, current) => (current.row > latest.row ? current : latest), potentialMatches[0]);
          if (existingRowInfoToUpdate) targetSheetRowForUpdate = existingRowInfoToUpdate.row;
      }

      let rowDataForSheet = new Array(TOTAL_COLUMNS_IN_APP_SHEET).fill(""); // From Config.gs

      if (targetSheetRowForUpdate !== -1 && existingRowInfoToUpdate) { 
        const currentSheetValues = dataSheet.getRange(targetSheetRowForUpdate, 1, 1, TOTAL_COLUMNS_IN_APP_SHEET).getValues()[0];
        rowDataForSheet = [...currentSheetValues]; 
        rowDataForSheet[PROCESSED_TIMESTAMP_COL-1] = currentTimestamp;
        const esDate = rowDataForSheet[EMAIL_DATE_COL-1]; if(!(esDate instanceof Date)||emailDate.getTime()>new Date(esDate).getTime())rowDataForSheet[EMAIL_DATE_COL-1]=emailDate;
        const elDate = rowDataForSheet[LAST_UPDATE_DATE_COL-1]; if(!(elDate instanceof Date)||emailDate.getTime()>new Date(elDate).getTime())rowDataForSheet[LAST_UPDATE_DATE_COL-1]=emailDate;
        rowDataForSheet[EMAIL_SUBJECT_COL-1]=emailSubject; rowDataForSheet[EMAIL_LINK_COL-1]=emailPermaLink; rowDataForSheet[EMAIL_ID_COL-1]=msgId; rowDataForSheet[PLATFORM_COL-1]=detectedPlatform;
        if(companyName!==MANUAL_REVIEW_NEEDED && (rowDataForSheet[COMPANY_COL-1]===MANUAL_REVIEW_NEEDED||companyName.toLowerCase()!==String(rowDataForSheet[COMPANY_COL-1]).toLowerCase()))rowDataForSheet[COMPANY_COL-1]=companyName;
        if(jobTitle!==MANUAL_REVIEW_NEEDED && (rowDataForSheet[JOB_TITLE_COL-1]===MANUAL_REVIEW_NEEDED||jobTitle.toLowerCase()!==String(rowDataForSheet[JOB_TITLE_COL-1]).toLowerCase()))rowDataForSheet[JOB_TITLE_COL-1]=jobTitle;
        const statInSheet=String(rowDataForSheet[STATUS_COL-1]).trim()||DEFAULT_STATUS;
        if(statInSheet!==ACCEPTED_STATUS||finalStatusToSet===ACCEPTED_STATUS){const curRank=STATUS_HIERARCHY[statInSheet]??0; const newRank=STATUS_HIERARCHY[finalStatusToSet]??0; if(newRank>=curRank||finalStatusToSet===REJECTED_STATUS||finalStatusToSet===OFFER_STATUS)rowDataForSheet[STATUS_COL-1]=finalStatusToSet;}
        const statAfterUpd=String(rowDataForSheet[STATUS_COL-1]); let peakStat=existingRowInfoToUpdate.peakStatus||String(rowDataForSheet[PEAK_STATUS_COL-1]).trim(); if(!peakStat||peakStat===MANUAL_REVIEW_NEEDED||peakStat==="")peakStat=DEFAULT_STATUS;
        const curPeakRank=STATUS_HIERARCHY[peakStat]??-2; const newStatRankPeak=STATUS_HIERARCHY[statAfterUpd]??-2; const exclPeak=new Set([REJECTED_STATUS,ACCEPTED_STATUS,MANUAL_REVIEW_NEEDED,"Update/Other"]);
        let updPeakVal=peakStat; if(newStatRankPeak>curPeakRank&&!exclPeak.has(statAfterUpd))updPeakVal=statAfterUpd; else if(peakStat===DEFAULT_STATUS&&!exclPeak.has(statAfterUpd)&&STATUS_HIERARCHY[statAfterUpd]>STATUS_HIERARCHY[DEFAULT_STATUS])updPeakVal=statAfterUpd;
        rowDataForSheet[PEAK_STATUS_COL-1]=updPeakVal;
        // NOTES_COL (index 11 for col 12) remains as is from currentSheetValues if not explicitly changed.
        dataSheet.getRange(targetSheetRowForUpdate, 1, 1, TOTAL_COLUMNS_IN_APP_SHEET).setValues([rowDataForSheet]);
        Logger.log(`[${FUNC_NAME} INFO] SHEET WRITE: Updated Row ${targetSheetRowForUpdate}. Status:"${statAfterUpd}", Peak:"${updPeakVal}"`);
        sheetUpdateSuccessCount++; sheetWriteOpSuccessThisMessage = true;
        const cKey= (rowDataForSheet[COMPANY_COL-1]!==MANUAL_REVIEW_NEEDED)?String(rowDataForSheet[COMPANY_COL-1]).toLowerCase():companyCacheKey;
        if(existingDataCache[cKey])existingDataCache[cKey]=existingDataCache[cKey].map(e=>e.row===targetSheetRowForUpdate?{...e,status:statAfterUpd,peakStatus:updPeakVal,emailId:msgId,title:rowDataForSheet[JOB_TITLE_COL-1]}:e);
      } else { 
        rowDataForSheet[PROCESSED_TIMESTAMP_COL-1]=currentTimestamp; rowDataForSheet[EMAIL_DATE_COL-1]=emailDate; rowDataForSheet[PLATFORM_COL-1]=detectedPlatform; rowDataForSheet[COMPANY_COL-1]=companyName; rowDataForSheet[JOB_TITLE_COL-1]=jobTitle; rowDataForSheet[STATUS_COL-1]=finalStatusToSet; rowDataForSheet[LAST_UPDATE_DATE_COL-1]=emailDate; rowDataForSheet[EMAIL_SUBJECT_COL-1]=emailSubject; rowDataForSheet[EMAIL_LINK_COL-1]=emailPermaLink; rowDataForSheet[EMAIL_ID_COL-1]=msgId;
        // NOTES_COL (index 11 for col 12) will be "" by default.
        const exclPeakInit=new Set([REJECTED_STATUS,ACCEPTED_STATUS,MANUAL_REVIEW_NEEDED,"Update/Other"]);
        if(!exclPeakInit.has(finalStatusToSet))rowDataForSheet[PEAK_STATUS_COL-1]=finalStatusToSet; else rowDataForSheet[PEAK_STATUS_COL-1]=DEFAULT_STATUS;
        dataSheet.appendRow(rowDataForSheet);
        const newSRN=dataSheet.getLastRow();
        Logger.log(`[${FUNC_NAME} INFO] SHEET WRITE: Appended Row ${newSRN}. Status:"${finalStatusToSet}", Peak:"${rowDataForSheet[PEAK_STATUS_COL - 1]}"`);
        newEntryCount++; sheetWriteOpSuccessThisMessage = true;
        const nKey=(rowDataForSheet[COMPANY_COL-1]!==MANUAL_REVIEW_NEEDED)?String(rowDataForSheet[COMPANY_COL-1]).toLowerCase():companyCacheKey;
        if(!existingDataCache[nKey])existingDataCache[nKey]=[]; existingDataCache[nKey].push({row:newSRN,emailId:msgId,company:rowDataForSheet[COMPANY_COL-1],title:rowDataForSheet[JOB_TITLE_COL-1],status:rowDataForSheet[STATUS_COL-1],peakStatus:rowDataForSheet[PEAK_STATUS_COL-1]});
      }

      if (sheetWriteOpSuccessThisMessage) {
        processedThisRunCount++; processedEmailIds.add(msgId);
        let msgOutcome=(requiresManualReview||companyName===MANUAL_REVIEW_NEEDED||jobTitle===MANUAL_REVIEW_NEEDED)?'manual':'done';
        if(threadProcessingOutcomes[threadId]!=='manual')threadProcessingOutcomes[threadId]=msgOutcome;
        if(msgOutcome==='manual')threadProcessingOutcomes[threadId]='manual';
      } else { processingErrorCount++; threadProcessingOutcomes[threadId]='manual'; Logger.log(`[${FUNC_NAME} ERROR] Sheet Write Fail Msg ${msgId}. Thread ${threadId} marked manual.`);}
    } catch (eMsgProc) {
      Logger.log(`[${FUNC_NAME} FATAL ERROR] Proc Msg ${msgId}(Thr ${threadId}): ${eMsgProc.message}\nStack:${eMsgProc.stack}`);
      threadProcessingOutcomes[threadId]='manual'; processingErrorCount++;
    }
    if(DEBUG_MODE){ const msgProcTime=(new Date().getTime()-processingStartTimeMsg.getTime())/1000; Logger.log(`--- [${FUNC_NAME}] End Msg ${i+1}/${messagesToSort.length} --- Time:${msgProcTime}s ---`);} 
    Utilities.sleep(200 + Math.floor(Math.random() * 100)); 
  }

  // --- Apply Final Labels ---
  Logger.log(`\n[${FUNC_NAME} INFO] Loop done. Processed:${processedThisRunCount}, Updates:${sheetUpdateSuccessCount}, New:${newEntryCount}, Errors:${processingErrorCount}.`);
  if(Object.keys(threadProcessingOutcomes).length > 0) {
      if (DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Final Thread Outcomes: ${JSON.stringify(threadProcessingOutcomes)}`);
      applyFinalLabels(threadProcessingOutcomes, procLbl, processedLblObj, manualLblObj);
  } else { Logger.log(`[${FUNC_NAME} INFO] No threads to re-label.`); }
  
  // --- Update Dashboard ---
  try {
    Logger.log(`[${FUNC_NAME} INFO] Final dashboard update...`);
    if (typeof updateDashboardMetrics === "function") updateDashboardMetrics(ss.getSheetByName(DASHBOARD_TAB_NAME), ss.getSheetByName(HELPER_SHEET_NAME), dataSheet);
  } catch (e_dash_final) { Logger.log(`[${FUNC_NAME} ERROR] Final dashboard update: ${e_dash_final.message}`); }

  const SCRIPT_END_TIME_FINAL = new Date(); // Renamed to avoid conflict
  Logger.log(`\n==== ${FUNC_NAME} FINISHED (${SCRIPT_END_TIME_FINAL.toLocaleString()}) === Total Time: ${(SCRIPT_END_TIME_FINAL.getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}


// --- Auto-Reject Stale Applications Function ---
function markStaleApplicationsAsRejected() {
  const FUNC_NAME = "markStaleApplicationsAsRejected";
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== ${FUNC_NAME}: START (${SCRIPT_START_TIME.toLocaleString()}) ====`);
  
  const { spreadsheet: ss } = getOrCreateSpreadsheetAndSheet(); // From SheetUtils.gs
  if (!ss) {
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Main spreadsheet access failed. Aborting.`);
    return;
  }

  let dataSheet; // "Applications" sheet
  try {
    dataSheet = ss.getSheetByName(APP_TRACKER_SHEET_TAB_NAME); // From Config.gs
    if (!dataSheet) {
        Logger.log(`[${FUNC_NAME} FATAL ERROR] Tab "${APP_TRACKER_SHEET_TAB_NAME}" not found in "${ss.getName()}". Aborting.`);
        return;
    }
  } catch (e) {
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Accessing tab "${APP_TRACKER_SHEET_TAB_NAME}": ${e.message}. Aborting.`);
    return;
  }
  Logger.log(`[${FUNC_NAME} INFO] Using "${dataSheet.getName()}" in "${ss.getName()}".`);

  const headerRow = 1;
  if (dataSheet.getLastRow() <= headerRow) {
    Logger.log(`[${FUNC_NAME} INFO] No data rows in "${dataSheet.getName()}" for stale check.`);
    return;
  }

  const dataRange = dataSheet.getRange(headerRow + 1, 1, dataSheet.getLastRow() - headerRow, dataSheet.getLastColumn());
  const sheetValues = dataRange.getValues(); // This is a 2D array

  const currentDate = new Date();
  const staleThresholdDate = new Date();
  staleThresholdDate.setDate(currentDate.getDate() - (WEEKS_THRESHOLD * 7)); // From Config.gs
  Logger.log(`[${FUNC_NAME} INFO] Stale if Last Update < ${staleThresholdDate.toLocaleDateString()} (Threshold: ${WEEKS_THRESHOLD} weeks)`);

  let updatedApplicationsCount = 0;
  let changesMadeToSheetValues = false;

  for (let i = 0; i < sheetValues.length; i++) {
    const currentRowArray = sheetValues[i];
    const actualSheetRowNumber = i + headerRow + 1;
    const currentStatus = currentRowArray[STATUS_COL - 1] ? String(currentRowArray[STATUS_COL - 1]).trim() : "";
    const lastUpdateDateValue = currentRowArray[LAST_UPDATE_DATE_COL - 1];
    let lastUpdateDate;

    if (lastUpdateDateValue instanceof Date && !isNaN(lastUpdateDateValue)) { lastUpdateDate = lastUpdateDateValue; }
    else if (lastUpdateDateValue && typeof lastUpdateDateValue === 'string' && lastUpdateDateValue.trim() !== "") {
      const parsed = new Date(lastUpdateDateValue); if (!isNaN(parsed)) lastUpdateDate = parsed; else continue;
    } else { continue; }

    if (FINAL_STATUSES_FOR_STALE_CHECK.has(currentStatus) || !currentStatus || currentStatus === MANUAL_REVIEW_NEEDED) continue;
    if (lastUpdateDate.getTime() >= staleThresholdDate.getTime()) continue;
    
    Logger.log(`[${FUNC_NAME} INFO Row ${actualSheetRowNumber}] MARKING STALE: "${currentStatus}" -> "${REJECTED_STATUS}"`);
    sheetValues[i][STATUS_COL - 1] = REJECTED_STATUS;
    sheetValues[i][LAST_UPDATE_DATE_COL - 1] = currentDate;
    sheetValues[i][PROCESSED_TIMESTAMP_COL - 1] = currentDate;
    updatedApplicationsCount++;
    changesMadeToSheetValues = true;
  }

  if (updatedApplicationsCount > 0 && changesMadeToSheetValues) {
    Logger.log(`[${FUNC_NAME} INFO] Found ${updatedApplicationsCount} stale apps. Writing changes...`);
    try {
      dataRange.setValues(sheetValues);
      Logger.log(`[${FUNC_NAME} INFO] Updated ${updatedApplicationsCount} stale applications.`);
    } catch (e) { Logger.log(`[${FUNC_NAME} ERROR] Sheet write failed: ${e.toString()}`); }
  } else { Logger.log(`[${FUNC_NAME} INFO] No stale applications found needing update.`); }
  
  const SCRIPT_END_TIME = new Date();
  Logger.log(`==== ${FUNC_NAME}: END (${SCRIPT_END_TIME.toLocaleString()}) ==== Time: ${(SCRIPT_END_TIME.getTime() - SCRIPT_START_TIME.getTime()) / 1000}s ====`);
}

// --- onOpen Function ---
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menuName = CUSTOM_MENU_NAME || '⚙️ CareerSuite.AI Tools'; // CUSTOM_MENU_NAME from Config.gs
  const menu = ui.createMenu(menuName);

  menu.addItem('▶️ RUN FULL PROJECT SETUP', 'runFullProjectInitialSetup');
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Module Setups')
      .addItem('Setup: Job Application Tracker', 'initialSetup_LabelsAndSheet')
      .addItem('Setup: Job Leads Tracker', 'runInitialSetup_JobLeadsModule')); // Assumed in Leads_Main.gs
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Manual Processing')
      .addItem('📧 Process Application Emails', 'processJobApplicationEmails')
      .addItem('📬 Process Job Leads', 'processJobLeads') // Assumed in Leads_Main.gs
      .addItem('🗑️ Mark Stale Applications', 'markStaleApplicationsAsRejected'));
  menu.addSeparator();
  menu.addSubMenu(ui.createMenu('Admin & Config')
      .addItem('🔑 Set Gemini API Key', 'setSharedGeminiApiKey_UI') // Assumed in AdminUtils.gs
      .addItem('🔍 Show All User Properties', 'showAllUserProperties') // Assumed in AdminUtils.gs
      .addItem('🔩 TEMPORARY: Set Hardcoded API Key', 'TEMPORARY_manualSetSharedGeminiApiKey')); // Assumed in AdminUtils.gs
  menu.addToUi();

  // --- Automatic Initial Setup for New Copies ---
  // This section attempts to run the initial setup automatically when a user opens
  // their unique copy of the sheet (with its own copied script project) for the first time.
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const setupCompleteFlag = 'initialSetupDone_vCSAI_1'; // Use a unique, versioned flag

    if (!scriptProperties.getProperty(setupCompleteFlag)) {
      Logger.log(`[onOpen INFO] Flag "${setupCompleteFlag}" not found. Checking conditions for initial setup.`);

      // Ensure this isn't running in a limited authorization mode (e.g., when a trigger fires headless)
      // For onOpen, e.authMode is usually FULL or LIMITED for the user opening it. NONE is rare for onOpen.
      if (e && e.authMode && e.authMode !== ScriptApp.AuthMode.NONE) {
        const activeSS = SpreadsheetApp.getActiveSpreadsheet();
        const currentSheetId = activeSS.getId();

        // Safety check: Ensure this isn't the original template sheet.
        // This is a secondary check; the primary scenario is this code runs in a *copied* script project
        // where TEMPLATE_SHEET_ID might not even be relevant if Config.js isn't perfectly copied or is altered.
        // However, if TEMPLATE_SHEET_ID *is* defined (e.g. from Config.js) and matches, skip.
        let isTemplate = false;
        if (typeof TEMPLATE_SHEET_ID !== 'undefined' && TEMPLATE_SHEET_ID && TEMPLATE_SHEET_ID !== "" && currentSheetId === TEMPLATE_SHEET_ID) {
          isTemplate = true;
          Logger.log(`[onOpen INFO] Current sheet (ID: ${currentSheetId}) is the template. Automatic setup skipped.`);
        }

        if (!isTemplate) {
          Logger.log(`[onOpen INFO] New copy (Sheet ID: ${currentSheetId}). Attempting automatic initial setup.`);
          // Call the main setup function, passing the active spreadsheet.
          const setupResult = runFullProjectInitialSetup(activeSS);

          if (setupResult && setupResult.success) {
            scriptProperties.setProperty(setupCompleteFlag, 'true');
            Logger.log(`[onOpen INFO] Automatic initial setup successful. Flag "${setupCompleteFlag}" set.`);
            // Optionally, notify the user of success if appropriate (though setup might have its own alerts)
            // ui.alert('Setup Complete', 'The initial setup for your sheet has completed successfully!', ui.ButtonSet.OK);
          } else {
            Logger.log(`[onOpen WARN] Automatic initial setup may have failed or was skipped. Result: ${JSON.stringify(setupResult)}`);
            // Optionally, alert the user that setup needs to be run manually
            // ui.alert('Setup Incomplete', 'The automatic initial setup did not complete as expected. Please try running "RUN FULL PROJECT SETUP" from the "${menuName}" menu.', ui.ButtonSet.OK);
          }
        }
      } else {
        Logger.log(`[onOpen INFO] Setup skipped due to AuthMode: ${e ? e.authMode : 'N/A (no event object)'}. This is normal if opened in a restricted context.`);
      }
    } else {
      Logger.log(`[onOpen INFO] Flag "${setupCompleteFlag}" is set. Automatic initial setup already performed or skipped previously.`);
    }
  } catch (err) {
    Logger.log(`[onOpen ERROR] Error during automatic initial setup attempt: ${err.toString()}\nStack: ${err.stack}`);
    // Avoid breaking the onOpen for menu creation if auto-setup fails.
    // ui.alert('Error', 'An error occurred during the onOpen process. Some automated setup steps may not have run. Please check logs or try manual setup.', ui.ButtonSet.OK);
  }
}
