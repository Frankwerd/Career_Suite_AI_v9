// File: WebApp_Endpoints.gs
// Description: Handles authenticated GET requests for the CareerSuite.AI Web App.
// Version: 2.1 (OAuth2 Integrated, Final)

/**
 * Main entry point for all GET requests to the Web App.
 * Routes requests based on an 'action' parameter after validating authentication.
 * @param {object} e The event parameter from the GET request.
 * @return {GoogleAppsScript.Content.TextOutput | GoogleAppsScript.HTML.HtmlOutput} A JSON or HTML response.
 */
function doGet(e) {
  const FUNC_NAME = "WebApp_doGet";
  try {
    const action = e.parameter.action;
    Logger.log(`[${FUNC_NAME}] Received GET request. Action: "${action}". User: ${Session.getEffectiveUser().getEmail()}`);
    
    // The presence of an `action` parameter implies a legitimate API call from our extension.
    // The `forceAuth` parameter is for the initial auth landing page.
    if (action) {
      // --- Route to the correct function based on the 'action' parameter ---
      switch (action) {
        case 'getOrCreateSheet':
          return doGet_getOrCreateSheet(e);
        case 'getWeeklyApplicationData':
          return doGet_WeeklyApplicationData(e);
        default:
          return createJsonResponse({ success: false, error: `Unknown action: ${action}` });
      }
    } else if (e.parameter.forceAuth) {
        return createAuthLandingPage();
    } else {
        // This case handles a direct visit to the URL without parameters.
        return createAuthLandingPage(); // Show a friendly page instead of an error.
    }

  } catch (error) {
    Logger.log(`[${FUNC_NAME}] CRITICAL Unhandled Error in doGet: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({ success: false, error: `An unexpected server error occurred: ${error.message}` });
  }
}

function doPost(e) {
  try {
    // --- User & Action Identification ---
    // Get the authenticated user's email. This is crucial for logging and multi-user contexts.
    const userEmail = e && e.user ? e.user.email : Session.getActiveUser().getEmail();
    Logger.log(`WebApp_Endpoints: doPost received a request from user: ${userEmail}.`);

    // Determine the specific action requested by the extension via URL parameter.
    // e.g., ?action=setApiKey
    const action = e && e.parameter ? e.parameter.action : null;

    // --- Action Routing ---
    // Based on the 'action' parameter, we route to the appropriate logic block.

    // === ACTION: setApiKey ===
    // This action securely receives the user's Gemini API key from the extension
    // and saves it to their private UserProperties store.
    if (action === 'setApiKey') {
      Logger.log(`[WebApp] Routing to 'setApiKey' action for user ${userEmail}.`);
      
      const postData = JSON.parse(e.postData.contents);
      const apiKey = postData.apiKey;

      // Validate the received key to ensure it's not empty or invalid.
      if (!apiKey || typeof apiKey !== 'string' || apiKey.length < 35) {
        Logger.log(`[WebApp ERROR] 'setApiKey' failed: Invalid API key provided by user ${userEmail}.`);
        return ContentService.createTextOutput(JSON.stringify({
          status: 'error',
          message: 'Invalid or missing API key provided in the request.'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      // GEMINI_API_KEY_PROPERTY is the constant from Config.gs
      PropertiesService.getUserProperties().setProperty(GEMINI_API_KEY_PROPERTY, apiKey);

      Logger.log(`[WebApp SUCCESS] Successfully saved Gemini API key to UserProperties for user ${userEmail}.`);
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'API key was successfully saved to the backend.'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // === DEFAULT ACTION: getOrCreateSheet ===
    // If no specific action is provided, the default behavior is to get or create
    // the user's main Job Tracker spreadsheet. This is triggered by the "Manage Job Tracker" button.
    if (!action) {
      Logger.log(`[WebApp] Routing to default 'getOrCreateSheet' action for user ${userEmail}.`);

      const existingSheetId = PropertiesService.getUserProperties().getProperty('userMjmSheetId');
      if (existingSheetId) {
        try {
          const existingSheet = SpreadsheetApp.openById(existingSheetId);
          if (existingSheet) {
            Logger.log(`[WebApp SUCCESS] User ${userEmail} already has a sheet. ID=${existingSheetId}`);
            return ContentService.createTextOutput(JSON.stringify({
              status: 'success',
              message: 'Your CareerSuite.AI Data sheet already exists.',
              sheetId: existingSheetId,
              sheetUrl: existingSheet.getUrl(),
              sheetName: existingSheet.getName()
            })).setMimeType(ContentService.MimeType.JSON);
          }
        } catch (openErr) {
          Logger.log(`[WebApp WARN] Stored sheet ID ${existingSheetId} for ${userEmail} was inaccessible: ${openErr.message}. Clearing property and creating a new sheet.`);
          PropertiesService.getUserProperties().deleteProperty('userMjmSheetId');
        }
      }

      // Logic to create a new sheet from the template
      const templateIdToUse = TEMPLATE_MJM_SHEET_ID; // From Config.gs
      if (!templateIdToUse || templateIdToUse.length < 20) {
        throw new Error("Server configuration error: Master Template Sheet ID is not set correctly in Config.gs.");
      }

      Logger.log(`[WebApp INFO] Creating new sheet from template ID ${templateIdToUse} for user ${userEmail}.`);
      const originalFile = DriveApp.getFileById(templateIdToUse);
      const newFileName = `CareerSuite.AI Data`;
      
      const newSpreadsheetFile = originalFile.makeCopy(newFileName);
      const newSpreadsheetId = newSpreadsheetFile.getId();
      
      newSpreadsheetFile.setOwner(userEmail);

      Logger.log(`[WebApp SUCCESS] New sheet created for ${userEmail}: "${newFileName}", ID=${newSpreadsheetId}. Ownership transferred.`);

      PropertiesService.getUserProperties().setProperty('userMjmSheetId', newSpreadsheetId);
      runFullProjectInitialSetup(SpreadsheetApp.openById(newSpreadsheetId));
      
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Your CareerSuite.AI Data sheet has been created and set up successfully!',
        sheetId: newSpreadsheetId,
        sheetUrl: newSpreadsheetFile.getUrl(),
        sheetName: newFileName
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Fallback for an unknown action parameter.
    Logger.log(`[WebApp WARN] An unknown POST action was requested: "${action}".`);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Unknown or unsupported POST action requested.'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Global Error Handling for any unexpected errors.
    Logger.log(`[WebApp CRITICAL ERROR] Error in doPost (Outer Catch): ${error.toString()}\nStack: ${error.stack}`);
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Failed to complete the requested action due to a server-side error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Creates a generic HTML response for the OAuth landing page.
 * @return {GoogleAppsScript.HTML.HtmlOutput}
 */
function createAuthLandingPage() {
    const htmlOutput = `
      <!DOCTYPE html><html><head><title>CareerSuite.AI Authorization</title>
      <style>body { font-family: sans-serif; margin: 20px; background-color: #f0f4f8; color: #333; text-align: center; } .container { background-color: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); display: inline-block; } h1 { color: #33658A; }</style></head><body><div class="container">
        <h1>CareerSuite.AI</h1><p>Authorization successful!</p><p>You can now close this tab and return to the extension.</p>
      </div></body></html>`;
    return HtmlService.createHtmlOutput(htmlOutput)
      .setTitle("CareerSuite.AI Authorization")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Creates a standardized JSON response object.
 * @param {object} payload - The JSON payload to send.
 * @return {GoogleAppsScript.Content.TextOutput}
 */
function createJsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles the logic for getting an existing sheet or creating a new one for the user.
 * This is the primary endpoint for the extension's "Manage Job Tracker" button.
 * @param {object} e The event parameter from the GET request.
 * @return {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function doGet_getOrCreateSheet(e) {
  const FUNC_NAME = "doGet_getOrCreateSheet";
  const userEmail = Session.getEffectiveUser().getEmail();
  const userProps = PropertiesService.getUserProperties();
  const existingSheetId = userProps.getProperty('userMjmSheetId');
  
  // 1. Check if a valid Sheet ID is already stored for the user.
  if (existingSheetId) {
    try {
      const existingSheet = SpreadsheetApp.openById(existingSheetId);
      Logger.log(`[${FUNC_NAME}] Found existing, valid sheet for ${userEmail}: ID=${existingSheetId}`);
      return createJsonResponse({
        status: "success",
        message: "Sheet already exists.",
        sheetId: existingSheetId,
        sheetUrl: existingSheet.getUrl()
      });
    } catch (openErr) {
      Logger.log(`[${FUNC_NAME}] Stored sheet ID ${existingSheetId} was invalid or inaccessible. Clearing property and creating a new sheet. Error: ${openErr.message}`);
      userProps.deleteProperty('userMjmSheetId');
    }
  }

  // 2. If no valid ID, create a new sheet from the template.
  Logger.log(`[${FUNC_NAME}] No valid sheet found for ${userEmail}. Creating from template...`);
  const templateId = TEMPLATE_SHEET_ID; // From Config.gs
  if (!templateId || templateId.length < 20) {
    return createJsonResponse({ status: 'error', message: 'Server configuration error: Master Template Sheet ID is invalid.' });
  }

  const templateFile = DriveApp.getFileById(templateId);
  const newFileName = `CareerSuite.AI Data`; // From Config.gs -> TARGET_SPREADSHEET_FILENAME
  const newSheetFile = templateFile.makeCopy(newFileName);
  const newSheetId = newSheetFile.getId();
  
  // 3. Run the full project setup on the newly created sheet.
  const newSheet = SpreadsheetApp.openById(newSheetId);
  const setupResult = runFullProjectInitialSetup(newSheet); // From Main.gs
  
  if (!setupResult.success) {
    Logger.log(`[${FUNC_NAME}] CRITICAL: Initial setup failed on new sheet ${newSheetId}.`);
    // Consider deleting the failed sheet to avoid clutter: DriveApp.getFileById(newSheetId).setTrashed(true);
    return createJsonResponse({
      status: "error",
      message: `Failed to initialize the new sheet. Please check script logs for details. Error: ${setupResult.message}`
    });
  }

  // 4. Save the new, successfully initialized sheet ID to user properties.
  userProps.setProperty('userMjmSheetId', newSheetId);
  Logger.log(`[${FUNC_NAME}] New sheet created and initialized for ${userEmail}. ID: ${newSheetId}`);

  return createJsonResponse({
    status: "success",
    message: "Your CareerSuite.AI Data sheet was created and set up successfully!",
    sheetId: newSheetId,
    sheetUrl: newSheet.getUrl()
  });
}

/**
 * Handles GET requests for aggregated weekly application data.
 * @param {object} e The event parameter from the GET request.
 * @return {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
function doGet_WeeklyApplicationData(e) {
  const FUNC_NAME = "doGet_WeeklyApplicationData";
  try {
    const userMjmSheetId = PropertiesService.getUserProperties().getProperty('userMjmSheetId');
    if (!userMjmSheetId) {
      return createJsonResponse({ 
          success: false, 
          error: "CareerSuite.AI Sheet ID not found. Please complete setup via the extension." 
      });
    }

    let ss;
    try {
        ss = SpreadsheetApp.openById(userMjmSheetId);
    } catch (sheetOpenErr) {
        Logger.log(`[${FUNC_NAME} ERROR] Error opening sheet ID ${userMjmSheetId}: ${sheetOpenErr.message}`);
        PropertiesService.getUserProperties().deleteProperty('userMjmSheetId');
        return createJsonResponse({ 
            success: false, 
            error: `Your saved Sheet ID is no longer accessible. Please re-link your sheet.`
        });
    }
    
    const helperSheet = ss.getSheetByName(HELPER_SHEET_NAME);
    if (!helperSheet) {
      return createJsonResponse({ 
          success: false, 
          error: `Helper data sheet ("${HELPER_SHEET_NAME}") not found. Please run 'Update Dashboard Metrics' from the tools menu in your sheet.`
      });
    }

    const headersRange = helperSheet.getRange("D1:E1").getDisplayValues();
    if (headersRange[0][0] !== "Week Starting" || headersRange[0][1] !== "Applications") {
        return createJsonResponse({ 
            success: false, 
            error: "Helper data format for weekly applications is incorrect."
        });
    }
    
    const lastDataRowInColD = helperSheet.getRange("D1:D").getValues().filter(String).length;
    let weeklyData = [];

    if (lastDataRowInColD > 1) {
        const maxWeeksToShow = 12; // Show up to 12 weeks of data
        const startRowForFetch = Math.max(2, lastDataRowInColD - maxWeeksToShow + 1);
        const numRowsToFetchActual = lastDataRowInColD - startRowForFetch + 1;
        
        if (numRowsToFetchActual > 0) {
            const rangeDataValues = helperSheet.getRange(startRowForFetch, 4, numRowsToFetchActual, 2).getDisplayValues();
            rangeDataValues.forEach(row => {
                if (row[0] && row[1]) {
                     weeklyData.push({ weekStarting: row[0], applications: row[1] });
                }
            });
        }
    }
    
    Logger.log(`[${FUNC_NAME} INFO] Fetched ${weeklyData.length} weekly data points.`);
    
    return createJsonResponse({ success: true, data: weeklyData });

  } catch (error) {
    Logger.log(`[${FUNC_NAME} ERROR] Error: ${error.toString()}\nStack: ${error.stack}`);
    return createJsonResponse({ 
        success: false, 
        error: `Error fetching weekly application data: ${error.toString()}`
    });
  }
}
