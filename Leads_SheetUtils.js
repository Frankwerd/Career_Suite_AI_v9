// File: Leads_SheetUtils.gs
// Handles GET and POST requests for the CareerSuite.AI Web App.

// NOTE: The hardcoded template ID has been removed from this file.
// The script now correctly uses the global TEMPLATE_SHEET_ID constant from Config.gs.

function doGet(e) {
  try {
    // This log helps confirm the function is triggered.
    Logger.log(`WebApp doGet triggered. This is typically the landing page after authorization.`);

    if (e && e.parameter && e.parameter.action) {
      // Keep this routing for future-proofing your GET requests
        const action = e.parameter.action;
        Logger.log(`WebApp doGet: Routing action "${action}"`);
        if (action === "getDashboardData") {
            return doGet_DashboardData(e);
        } else if (action === "getWeeklyApplicationData") {
            return doGet_WeeklyApplicationData(e);
        }
    }
    
    // This static HTML is now safer as it doesn't depend on userEmail.
    let htmlOutput = `
      <!DOCTYPE html>
      <html>
        <head>
          <title>CareerSuite.AI Setup & Authorization</title>
          <style>
            body { font-family: Arial, sans-serif; margin: 20px; background-color: #f0f4f8; color: #333; text-align: center; }
            .container { background-color: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); display: inline-block; max-width: 600px; }
            h1 { color: #33658A; }
            p { font-size: 1.1em; margin-bottom: 15px; line-height: 1.6; }
            .note { font-size: 0.9em; color: #555; margin-top: 25px; }
          </style>
        </head>
        <body>
          <div class="container">
            <h1><img src="https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_74x24dp.png" alt="Google Icon" style="vertical-align: middle; height: 24px; margin-right: 8px;">CareerSuite.AI</h1>
            <p>Authorization successful!</p>
            <p>You can now close this tab and return to the CareerSuite.AI extension.</p>
            <p>Click the 'Manage Job Tracker' button again to complete your sheet setup.</p>
            <p class="note">This Web App is part of the CareerSuite.AI Chrome Extension setup process.</p>
          </div>
        </body>
      </html>`;
    
    return HtmlService.createHtmlOutput(htmlOutput)
      .setTitle("CareerSuite.AI Authorization")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log("WebApp_Endpoints: Error in doGet (OAuth/Landing): " + error.toString() + "\nStack: " + error.stack);
    return HtmlService.createHtmlOutput(
      `<h1>Error</h1><p>An unexpected error occurred: ${error.toString()}</p>`
    );
  }
}

/**
 * Handles POST requests, primarily for initializing the user's sheet.
 */
function doPost(e) {
  try {
    const userEmail = e && e.user ? e.user.email : Session.getActiveUser().getEmail();
    const activeUserKey = Session.getEffectiveUser().getEmail();
    
    Logger.log(`WebApp_Endpoints: doPost called. User interacting: ${userEmail}, Effective user (for quota): ${activeUserKey}.`);
    
    const action = e && e.parameter ? e.parameter.action : null;

    // Default action is to get or create the sheet
    if (!action) {
      const existingSheetId = PropertiesService.getUserProperties().getProperty('userMjmSheetId');
      if (existingSheetId) {
        try {
          const existingSheet = SpreadsheetApp.openById(existingSheetId);
          if (existingSheet) {
            Logger.log(`Sheet already exists for ${userEmail}: ID=${existingSheetId}`);
            return ContentService.createTextOutput(JSON.stringify({ 
                status: "success", 
                message: "Your CareerSuite.AI Data sheet already exists.",
                sheetId: existingSheetId,
                sheetUrl: existingSheet.getUrl(),
                sheetName: existingSheet.getName()
            })).setMimeType(ContentService.MimeType.JSON);
          }
        } catch (openErr) {
            Logger.log(`Stored sheet ID ${existingSheetId} for ${userEmail} was inaccessible: ${openErr.message}. Clearing property.`);
            PropertiesService.getUserProperties().deleteProperty('userMjmSheetId');
        }
      }
      
      const templateIdToUse = TEMPLATE_SHEET_ID; // Now uses the global constant from Config.js
      if (!templateIdToUse || templateIdToUse.length < 20) { // Simple validity check
          throw new Error("Server configuration error: Master Template Sheet ID is not set correctly in Config.gs.");
      }
      Logger.log(`Using Template ID: ${templateIdToUse} for user ${userEmail}`);

      const originalFile = DriveApp.getFileById(templateIdToUse);
      const newFileName = `CareerSuite.AI Data`;
      
      const newSpreadsheetFile = originalFile.makeCopy(newFileName);
      const newSpreadsheetId = newSpreadsheetFile.getId();
      const newSpreadsheetUrl = newSpreadsheetFile.getUrl();
      Logger.log(`New sheet created for ${userEmail}: "${newFileName}", ID=${newSpreadsheetId}`);

      PropertiesService.getUserProperties().setProperty('userMjmSheetId', newSpreadsheetId);

      return ContentService.createTextOutput(JSON.stringify({ 
          status: "success", 
          message: "Your CareerSuite.AI Data sheet has been created successfully!",
          sheetId: newSpreadsheetId,
          sheetUrl: newSpreadsheetUrl,
          sheetName: newFileName
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Route to other potential POST actions in the future if needed
    // else if (action === "someOtherAction") { ... }

    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown POST action requested.' })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("WebApp_Endpoints: Error in doPost (Outer Catch): " + error.toString() + "\nStack: " + error.stack);
    return ContentService.createTextOutput(JSON.stringify({ 
        status: "error", 
        message: "Failed to complete sheet setup due to a server error: " + error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Handles GET requests for dashboard data.
 * Called via ?action=getDashboardData
 */
function doGet_DashboardData(e) {
  try {
    const userEmail = Session.getActiveUser().getEmail(); // For logging
    console.log(`WebApp_Endpoints: doGet_DashboardData called by ${userEmail}.`);

    const userMjmSheetId = PropertiesService.getUserProperties().getProperty('userMjmSheetId');
    if (!userMjmSheetId) {
      console.warn(`doGet_DashboardData: No userMjmSheetId found for ${userEmail}.`);
      return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          error: "CareerSuite.AI Sheet ID not found. Please complete setup via the extension." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let ss;
    try {
        ss = SpreadsheetApp.openById(userMjmSheetId);
    } catch (sheetOpenErr) {
        console.error(`doGet_DashboardData: Error opening sheet ID ${userMjmSheetId} for ${userEmail}: ${sheetOpenErr.message}`);
        PropertiesService.getUserProperties().deleteProperty('userMjmSheetId'); // Clear invalid ID
        return ContentService.createTextOutput(JSON.stringify({ 
            success: false, 
            error: `Your saved Sheet ID (${userMjmSheetId.substring(0,10)}...) is no longer accessible. Please re-create or re-link your sheet via the extension. Error: ${sheetOpenErr.message}`
        })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const dashboardSheet = ss.getSheetByName(DASHBOARD_TAB_NAME); // DASHBOARD_TAB_NAME from Config.gs

    if (!dashboardSheet) {
      console.warn(`doGet_DashboardData: Dashboard sheet named "${DASHBOARD_TAB_NAME}" not found in tracker for ${userEmail} (Sheet ID: ${userMjmSheetId}).`);
      return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          error: `Dashboard sheet ("${DASHBOARD_TAB_NAME}") not found in your CareerSuite.AI Data sheet. Setup may be incomplete.`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Read values directly from the scorecard cells on the "Dashboard" sheet
    const totalApplications = dashboardSheet.getRange("C5").getDisplayValue(); // Use getDisplayValue for formatted numbers/percentages
    const activeApplications = dashboardSheet.getRange("C7").getDisplayValue();
    const currentlyInterviewing = dashboardSheet.getRange("I7").getDisplayValue();
    // Add more metrics if needed, e.g.,
    // const peakInterviews = dashboardSheet.getRange("F5").getDisplayValue();
    // const peakOffers = dashboardSheet.getRange("F7").getDisplayValue();
    // const interviewRate = dashboardSheet.getRange("I5").getDisplayValue(); // This will be a string like "50.00%"

    const data = {
      totalApplications: totalApplications, // Will be string, extension can parse if needed
      activeApplications: activeApplications,
      interviewing: currentlyInterviewing
      // peakInterviews: peakInterviews,
      // peakOffers: peakOffers,
      // interviewRate: interviewRate 
    };
    console.log(`doGet_DashboardData: Successfully fetched data for ${userEmail}: ${JSON.stringify(data)}`);
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("WebApp_Endpoints: Error in doGet_DashboardData: " + error.toString() + "\nStack: " + error.stack);
    return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: "An error occurred while fetching dashboard data: " + error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Simple XML Escaper for HTML Service output
function escapeXmlSimple(text) {
    if (text === null || typeof text === 'undefined') return '';
    return String(text)
        .replace(/&/g, "&")
        .replace(/</g, "<")
        .replace(/>/g, ">")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "'");
    }


// For testing template access directly from Apps Script editor
function manualTestTemplateAccess() {
  const id = TEMPLATE_SHEET_ID; // Uses the global constant from Config.js
  if (!id || id === "YOUR_MASTER_TEMPLATE_SHEET_ID_GOES_HERE") {
    Logger.log("manualTestTemplateAccess: TEMPLATE_SHEET_ID is not set in Config.gs. Please set it and try again.");
    SpreadsheetApp.getUi().alert("Template ID Not Set", "Please set the TEMPLATE_SHEET_ID constant in Config.gs before running this test.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  try {
    const file = DriveApp.getFileById(id);
    Logger.log("Template found: " + file.getName() + " (ID: " + file.getId() + ")");
    const copy = file.makeCopy("MANUAL TEST COPY - DELETE ME - " + new Date().toLocaleTimeString());
    Logger.log("Test copy created: " + copy.getName() + ", ID: " + copy.getId());
    // DriveApp.getFileById(copy.getId()).setTrashed(true); // Clean up immediately
    // Logger.log("Test copy trashed. You may need to manually empty trash in Drive if you want it fully gone.");
    SpreadsheetApp.getUi().alert("Test Successful", `Template found: ${file.getName()}\nTest copy created: ${copy.getName()} (ID: ${copy.getId()})\n\nPLEASE DELETE THE TEST COPY MANUALLY FROM YOUR GOOGLE DRIVE.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log("Error accessing/copying template: " + e.toString());
    SpreadsheetApp.getUi().alert("Test Failed", "Error accessing or copying template: " + e.toString() + "\n\nEnsure TEMPLATE_SHEET_ID in Config.gs is correct and you have access to the template file.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Handles GET requests for aggregated weekly application data.
 * Called via ?action=getWeeklyApplicationData
 * Reads from DashboardHelperData sheet, columns D (Week Starting) and E (Applications).
 */
function doGet_WeeklyApplicationData(e) {
  const FUNC_NAME = "doGet_WeeklyApplicationData";
  try {
    const userMjmSheetId = PropertiesService.getUserProperties().getProperty('userMjmSheetId');
    if (!userMjmSheetId) {
      Logger.log(`[${FUNC_NAME} WARN] No userMjmSheetId found.`);
      return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          error: "CareerSuite.AI Sheet ID not found. Please complete setup via the extension." 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    let ss;
    try {
        ss = SpreadsheetApp.openById(userMjmSheetId);
    } catch (sheetOpenErr) {
        Logger.log(`[${FUNC_NAME} ERROR] Error opening sheet ID ${userMjmSheetId}: ${sheetOpenErr.message}`);
        PropertiesService.getUserProperties().deleteProperty('userMjmSheetId'); // Clear invalid ID
        return ContentService.createTextOutput(JSON.stringify({ 
            success: false, 
            error: `Your saved Sheet ID is no longer accessible. Please re-create or re-link. Error: ${sheetOpenErr.message}`
        })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const helperSheet = ss.getSheetByName(HELPER_SHEET_NAME); // From Config.gs
    if (!helperSheet) {
      Logger.log(`[${FUNC_NAME} WARN] Helper sheet "${HELPER_SHEET_NAME}" not found.`);
      return ContentService.createTextOutput(JSON.stringify({ 
          success: false, 
          error: `Helper data sheet ("${HELPER_SHEET_NAME}") not found. Setup may be incomplete or run 'Update Dashboard Metrics'.`
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const headersRange = helperSheet.getRange("D1:E1").getDisplayValues();
    if (headersRange[0][0] !== "Week Starting" || headersRange[0][1] !== "Applications") {
        Logger.log(`[${FUNC_NAME} WARN] Expected headers "Week Starting" or "Applications" not found in D1 or E1 of helper sheet. Current headers: D1='${headersRange[0][0]}', E1='${headersRange[0][1]}'`);
        return ContentService.createTextOutput(JSON.stringify({ 
            success: false, 
            error: "Helper data format for weekly applications is incorrect. Please ensure dashboard setup/update has run."
        })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const lastDataRowInColD = helperSheet.getRange("D1:D").getValues().filter(String).length;
    let weeklyData = [];

    if (lastDataRowInColD > 1) { 
        const maxWeeksToShow = 8;
        const startRowForFetch = Math.max(2, lastDataRowInColD - maxWeeksToShow + 1); 
        const numRowsToFetchActual = lastDataRowInColD - startRowForFetch + 1;
        
        if (numRowsToFetchActual > 0) {
            const rangeDataValues = helperSheet.getRange(startRowForFetch, 4, numRowsToFetchActual, 2).getDisplayValues(); // Columns D:E
            rangeDataValues.forEach(row => {
                if (row[0] && row[1]) { 
                     weeklyData.push({ weekStarting: row[0], applications: row[1] });
                }
            });
        }
    }
    
    Logger.log(`[${FUNC_NAME} INFO] Fetched ${weeklyData.length} weekly data points. Last helper data row in D: ${lastDataRowInColD}`);
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: weeklyData }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(`[${FUNC_NAME} ERROR] Error: ${error.toString()}\nStack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: `Error fetching weekly application data: ${error.toString()}`
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
