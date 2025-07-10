// File: Leads_Main.gs
// Description: Contains the primary functions for the Job Leads Tracker module,
// including initial setup of the leads sheet/labels/filters and the
// ongoing processing of job lead emails.

/**
 * Sets up the Job Leads Tracker module:
 * - Ensures the "Potential Job Leads" sheet exists and formats it using SheetUtils.
 * - Creates necessary Gmail labels and filter.
 * - Sets up a daily trigger for processing new job leads.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} passedSpreadsheet The spreadsheet object to set up.
 * @return {{success: boolean, messages: Array<string>}} Result of the setup.
 */
function runInitialSetup_JobLeadsModule(passedSpreadsheet) {
  const FUNC_NAME = "runInitialSetup_JobLeadsModule";
  Logger.log(`\n==== ${FUNC_NAME}: STARTING - Leads Module Setup ====`);
  let setupMessages = [];
  let leadsModuleSetupSuccess = true;
  let activeSSLeads;

  if (passedSpreadsheet && typeof passedSpreadsheet.getId === 'function') {
    activeSSLeads = passedSpreadsheet;
    // Logger.log(`[${FUNC_NAME} INFO] Using PASSED spreadsheet ID: ${activeSSLeads.getId()}.`); // Redundant with Main.gs log
  } else {
    Logger.log(`[${FUNC_NAME} WARN] No spreadsheet passed. Fallback get/create...`);
    const { spreadsheet: mainSpreadsheet } = getOrCreateSpreadsheetAndSheet(); // From SheetUtils.gs
    activeSSLeads = mainSpreadsheet;
  }
  
  if (!activeSSLeads) {
    const errMsg = `CRITICAL [${FUNC_NAME}]: Could not obtain valid spreadsheet.`;
    Logger.log(errMsg); return { success: false, messages: [errMsg] };
  }
  setupMessages.push(`Using Spreadsheet: "${activeSSLeads.getName()}".`);

  try {
    // --- Step 1: Get/Create Leads Sheet Tab & Format ---
    Logger.log(`[${FUNC_NAME} INFO] Setting up sheet: "${LEADS_SHEET_TAB_NAME}"`);
    let leadsSheet = activeSSLeads.getSheetByName(LEADS_SHEET_TAB_NAME);
    if (!leadsSheet) {
      leadsSheet = activeSSLeads.insertSheet(LEADS_SHEET_TAB_NAME);
      Logger.log(`[${FUNC_NAME} INFO] CREATED new tab: "${LEADS_SHEET_TAB_NAME}".`);
    } else { Logger.log(`[${FUNC_NAME} INFO] Found EXISTING tab: "${LEADS_SHEET_TAB_NAME}".`); }
    
    if (!leadsSheet) throw new Error(`Get/Create FAILED for sheet: "${LEADS_SHEET_TAB_NAME}".`);
    
    // Call the generic formatting function from SheetUtils.gs
    // LEADS_SHEET_COLUMN_WIDTHS and LEADS_SHEET_HEADERS are from Config.gs
    if (!setupSheetFormatting(leadsSheet, 
                          LEADS_SHEET_HEADERS, 
                          LEADS_SHEET_COLUMN_WIDTHS,
                          true, // applyBandingFlag = true
                          SpreadsheetApp.BandingTheme.YELLOW // <<< SPECIFY YELLOW THEME (or CYAN, GREEN, GREY if yellow isn't distinct enough)
                         )) { /* throw error */ }
    leadsSheet.setTabColor(BRAND_COLORS.HUNYADI_YELLOW); // From Config.gs
    setupMessages.push(`Sheet "${LEADS_SHEET_TAB_NAME}": Setup OK. Color: Hunyadi Yellow.`);

    // --- Step 2: Gmail Label and Filter Setup ---
    Logger.log(`[${FUNC_NAME} INFO] Setting up Gmail labels & filters for Leads...`);
    getOrCreateLabel(MASTER_GMAIL_LABEL_PARENT); Utilities.sleep(100);
    getOrCreateLabel(LEADS_GMAIL_LABEL_PARENT); Utilities.sleep(100);
    const needsProcessLabelObject = getOrCreateLabel(LEADS_GMAIL_LABEL_TO_PROCESS); Utilities.sleep(100);
    const doneProcessLabelObject = getOrCreateLabel(LEADS_GMAIL_LABEL_PROCESSED); Utilities.sleep(100);
    if (!needsProcessLabelObject || !doneProcessLabelObject) throw new Error("Failed to create/verify core Leads Gmail labels.");
    
    let needsProcessLeadLabelId = null;
    const advGmailService = Gmail; // Assumes Advanced Gmail API Service is enabled
    Utilities.sleep(300);
    const labelsListRespLeads = advGmailService.Users.Labels.list('me');
    if (labelsListRespLeads.labels && labelsListRespLeads.labels.length > 0) {
        const targetLabelInfoLeads = labelsListRespLeads.labels.find(l => l.name === LEADS_GMAIL_LABEL_TO_PROCESS);
        if (targetLabelInfoLeads && targetLabelInfoLeads.id) needsProcessLeadLabelId = targetLabelInfoLeads.id;
    }
    if (!needsProcessLeadLabelId) throw new Error(`CRITICAL: Could not get ID for Gmail label "${LEADS_GMAIL_LABEL_TO_PROCESS}".`);
    setupMessages.push(`Leads Labels & 'To Process' ID: OK (${needsProcessLeadLabelId}).`);

    const filterQueryLeadsConst = LEADS_GMAIL_FILTER_QUERY; // from Config.gs
    let leadsFilterExists = false;
    const existingLeadsFiltersResponse = advGmailService.Users.Settings.Filters.list('me'); // Call it once

    // Robust check for the response structure
    if (existingLeadsFiltersResponse && existingLeadsFiltersResponse.filter && Array.isArray(existingLeadsFiltersResponse.filter)) {
        leadsFilterExists = existingLeadsFiltersResponse.filter.some(f => 
            f.criteria?.query === filterQueryLeadsConst && f.action?.addLabelIds?.includes(needsProcessLeadLabelId));
    } else if (existingLeadsFiltersResponse && !existingLeadsFiltersResponse.hasOwnProperty('filter')) {
        Logger.log(`[${FUNC_NAME} INFO] No 'filter' property in Gmail response (user likely has no filters). Assuming filter does not exist.`);
        leadsFilterExists = false; // Explicitly set
    } else {
        // This case handles if existingLeadsFiltersResponse is null or unexpected
        Logger.log(`[${FUNC_NAME} WARN] Unexpected response or null from Gmail Filters.list('me'). Assuming filter does not exist. Response: ${JSON.stringify(existingLeadsFiltersResponse)}`);
        leadsFilterExists = false; // Explicitly set
    }
    
    if (!leadsFilterExists) {
        const leadsFilterResource = { criteria: { query: filterQueryLeadsConst }, action: { addLabelIds: [needsProcessLeadLabelId], removeLabelIds: ['INBOX'] } };
        const createdLeadsFilter = advGmailService.Users.Settings.Filters.create(leadsFilterResource, 'me');
        if (!createdLeadsFilter || !createdLeadsFilter.id) {
             throw new Error(`Gmail filter creation for leads FAILED or no ID. Response: ${JSON.stringify(createdLeadsFilter)}`);
        }
        setupMessages.push("Leads Filter: CREATED.");
    } else { 
        setupMessages.push("Leads Filter: Exists."); 
    }
    
    // --- Step 3: Store Configuration (UserProperties for label IDs) ---
    const userPropsLeads = PropertiesService.getUserProperties();
    if (needsProcessLeadLabelId) userPropsLeads.setProperty(LEADS_USER_PROPERTY_TO_PROCESS_LABEL_ID, needsProcessLeadLabelId); else leadsModuleSetupSuccess = false;
    
    let doneProcessLeadLabelId = null;
    if(doneProcessLabelObject) { // Get ID for "Processed" label
        const doneLabelInfoLeads = labelsListRespLeads.labels?.find(l => l.name === LEADS_GMAIL_LABEL_PROCESSED);
        if (doneLabelInfoLeads?.id) {
             doneProcessLeadLabelId = doneLabelInfoLeads.id;
             userPropsLeads.setProperty(LEADS_USER_PROPERTY_PROCESSED_LABEL_ID, doneProcessLeadLabelId);
        }
    }
    if (!doneProcessLeadLabelId) leadsModuleSetupSuccess = false; // Critical if ID not found/stored for processing
    setupMessages.push(`UserProperties for Leads labels updated (ToProcessID: ${needsProcessLeadLabelId}, ProcessedID: ${doneProcessLeadLabelId}).`);

    // --- Step 4: Create Time-Driven Trigger ---
    if (createTimeDrivenTrigger('processJobLeads', 3)) { // Assumed in Triggers.gs, runs every 3 hours
        setupMessages.push("Trigger 'processJobLeads': CREATED.");
    } else { 
        setupMessages.push("Trigger 'processJobLeads': Exists/Verified."); 
    }

  } catch (e) {
    Logger.log(`[${FUNC_NAME} CRITICAL ERROR]: ${e.toString()}\nStack: ${e.stack || 'No stack'}`);
    setupMessages.push(`CRITICAL ERROR: ${e.message}`); leadsModuleSetupSuccess = false;
    try{ SpreadsheetApp.getUi().alert('Leads Module Setup Error', `Error: ${e.message}. Check logs.`, SpreadsheetApp.getUi().ButtonSet.OK); } catch(uiErr){ Logger.log("UI alert fail: " + uiErr.message);}
  }
  Logger.log(`Job Leads Module Setup: ${leadsModuleSetupSuccess ? "SUCCESSFUL" : "ISSUES"}.`);
  return { success: leadsModuleSetupSuccess, messages: setupMessages };
}

/**
 * Processes emails labeled for job leads.
 */
function processJobLeads() {
  const FUNC_NAME = "processJobLeads";
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== ${FUNC_NAME}: STARTING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  // --- 1. Configuration & Get Spreadsheet/Sheet ---
  const scriptProperties = PropertiesService.getScriptProperties(); // Changed from UserProperties
  const geminiApiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY); // From Config.gs

  if (geminiApiKey) {
    Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Retrieved key for "${GEMINI_API_KEY_PROPERTY}" from ScriptProperties. Value (masked): ${geminiApiKey.substring(0,4)}...${geminiApiKey.substring(geminiApiKey.length-4)}`);
    if (geminiApiKey.trim() !== "" && geminiApiKey.startsWith("AIza") && geminiApiKey.length > 30) {
        Logger.log(`[${FUNC_NAME} INFO] Gemini API Key (ScriptProperties) is valid for Leads processing.`);
    } else {
        Logger.log(`[${FUNC_NAME} WARN] Gemini API Key (ScriptProperties) found but INvalid. callGemini_forJobLeads might use mock data or fail if this key is passed without its own internal placeholder check.`);
        Logger.log(`[${FUNC_NAME} DEBUG_API_KEY] Reason: Key failed validation. Length: ${geminiApiKey.length}, StartsWith AIza: ${geminiApiKey.startsWith("AIza")}`);
        // The callGemini_forJobLeads function has a mock fallback if API key is placeholder-like
    }
  } else {
    Logger.log(`[${FUNC_NAME} WARN] Gemini API Key NOT FOUND in ScriptProperties for "${GEMINI_API_KEY_PROPERTY}". callGemini_forJobLeads will use mock/fail.`);
  }

  const { spreadsheet: activeSS } = getOrCreateSpreadsheetAndSheet(); // From SheetUtils.gs
  if (!activeSS) {
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Main spreadsheet could not be determined. Aborting.`);
    return;
  }

  const { sheet: leadsDataSheet, headerMap: leadsHeaderMap } = getSheetAndHeaderMapping_forLeads(activeSS.getId(), LEADS_SHEET_TAB_NAME); // From Leads_SheetUtils.gs
  if (!leadsDataSheet || !leadsHeaderMap || Object.keys(leadsHeaderMap).length === 0) { 
    Logger.log(`[${FUNC_NAME} FATAL ERROR] Leads sheet "${LEADS_SHEET_TAB_NAME}" or headers not found/mapped in SS ID ${activeSS.getId()}. Aborting.`); 
    return; 
  }
  Logger.log(`[${FUNC_NAME} INFO] Processing leads against: "${activeSS.getName()}", Leads Tab: "${leadsDataSheet.getName()}"`);

  // --- Get Gmail Labels ---
  const needsProcessLabelName = LEADS_GMAIL_LABEL_TO_PROCESS; // From Config.gs
  const doneProcessLabelName = LEADS_GMAIL_LABEL_PROCESSED;   // From Config.gs
  
  const needsProcessLabel = GmailApp.getUserLabelByName(needsProcessLabelName);
  const doneProcessLabel = GmailApp.getUserLabelByName(doneProcessLabelName); 
  if (!needsProcessLabel) { Logger.log(`[${FUNC_NAME} FATAL ERROR] Label "${needsProcessLabelName}" not found. Aborting.`); return; }
  if (!doneProcessLabel) { Logger.log(`[${FUNC_NAME} WARN] Label "${doneProcessLabelName}" not found. Processed leads will only be unlabelled from 'To Process'.`); }

  // --- Preload Processed Email IDs ---
  const processedLeadEmailIds = getProcessedEmailIdsFromSheet_forLeads(leadsDataSheet, leadsHeaderMap); // From Leads_SheetUtils.gs
  Logger.log(`[${FUNC_NAME} INFO] Preloaded ${processedLeadEmailIds.size} email IDs already processed for leads.`);

  // --- Fetch and Process Emails ---
  const LEADS_THREAD_LIMIT = 10; 
  const LEADS_MESSAGE_LIMIT_PER_RUN = 15;
  let messagesProcessedThisRunCounter = 0;
  const leadThreadsToProcess = needsProcessLabel.getThreads(0, LEADS_THREAD_LIMIT);
  Logger.log(`[${FUNC_NAME} INFO] Found ${leadThreadsToProcess.length} threads in "${needsProcessLabelName}".`);

  for (const thread of leadThreadsToProcess) {
    if (messagesProcessedThisRunCounter >= LEADS_MESSAGE_LIMIT_PER_RUN) { 
      Logger.log(`[${FUNC_NAME} INFO] Message processing limit (${LEADS_MESSAGE_LIMIT_PER_RUN}) reached for this run.`); 
      break; 
    }
    const scriptRunTimeSeconds = (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000;
    if (scriptRunTimeSeconds > 320) { // ~5min 20sec, leaving buffer for 6min limit
      Logger.log(`[${FUNC_NAME} WARN] Execution time limit (${scriptRunTimeSeconds}s) approaching. Stopping further thread processing.`); 
      break; 
    }

    const messagesInThread = thread.getMessages();
    let threadContainedAtLeastOneNewMessage = false;
    let allNewMessagesInThisThreadProcessedSuccessfully = true; // Assume success for new messages in this thread

    for (const message of messagesInThread) {
      if (messagesProcessedThisRunCounter >= LEADS_MESSAGE_LIMIT_PER_RUN) break; // Check limit per message too
      
      const msgId = message.getId();
      if (processedLeadEmailIds.has(msgId)) { // Check against preloaded IDs
        if (DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Msg ID ${msgId} in thread ${thread.getId()} already processed. Skipping.`);
        continue; 
      }
      
      threadContainedAtLeastOneNewMessage = true; 
      Logger.log(`\n--- [${FUNC_NAME}] Processing NEW Lead Msg ID: ${msgId}, Thread: ${thread.getId()}, Subject: "${message.getSubject()}" ---`);
      messagesProcessedThisRunCounter++;
      let currentMessageHandledNoErrors = false; 

      try {
        let emailBody = message.getPlainBody();
        if (typeof emailBody !== 'string' || emailBody.trim() === "") {
          Logger.log(`[${FUNC_NAME} WARN] Msg ${msgId}: Body is empty or not a string. Skipping AI call for this message.`);
          // Not writing an error to sheet, just skipping this specific message for AI processing.
          // If all other messages in thread are fine, thread can still be marked done.
          currentMessageHandledNoErrors = true; // Considered "handled" as there's nothing to parse
          processedLeadEmailIds.add(msgId); // Add to set to avoid re-processing this empty message
          continue; 
        }

        // geminiApiKey is passed to callGemini_forJobLeads, which might use mock data if key is invalid/placeholder
        const geminiApiResponse = callGemini_forJobLeads(emailBody, geminiApiKey); // From GeminiService.gs
        
        if (geminiApiResponse && geminiApiResponse.success) {
          const extractedJobsArray = parseGeminiResponse_forJobLeads(geminiApiResponse.data); // From GeminiService.gs
          if (extractedJobsArray && extractedJobsArray.length > 0) {
            Logger.log(`[${FUNC_NAME} INFO] Gemini extracted ${extractedJobsArray.length} job(s) from msg ${msgId}.`);
            let atLeastOneValidJobWrittenThisMessage = false;
            for (const jobData of extractedJobsArray) {
              if (jobData && jobData.jobTitle && String(jobData.jobTitle).toLowerCase() !== 'n/a' && String(jobData.jobTitle).toLowerCase() !== 'error') {
                jobData.dateAdded = message.getDate(); 
                jobData.sourceEmailSubject = message.getSubject().substring(0,500); // Keep this
                jobData.sourceEmailId = msgId; 
                jobData.status = "New"; 
                jobData.processedTimestamp = new Date();
                
                writeJobDataToSheet_forLeads(leadsDataSheet, jobData, leadsHeaderMap); // From Leads_SheetUtils.gs
                atLeastOneGoodJobWrittenThisMessage = true;
              } else { 
                if (DEBUG_MODE) Logger.log(`[${FUNC_NAME} DEBUG] Job from msg ${msgId} was N/A/error or missing title. Skipping sheet write: ${JSON.stringify(jobData)}`); 
              }
            }
            // If at least one job was extracted and written, consider this message "handled successfully" for now.
            if (atLeastOneValidJobWrittenThisMessage) currentMessageHandledNoErrors = true;
            else { 
              Logger.log(`[${FUNC_NAME} INFO] Msg ${msgId}: Gemini success, and parsed jobs, but no valid/writable jobs after filtering. Considered handled.`);
              currentMessageHandledNoErrors = true; // Gemini worked, parsing worked, just no "good" jobs.
            }
          } else { 
            Logger.log(`[${FUNC_NAME} INFO] Msg ${msgId}: Gemini API call success, but parsing response yielded no distinct job listings.`);
            currentMessageHandledNoErrors = true; // Gemini worked, nothing to parse = message handled.
          }
        } else { // Gemini API call itself failed
          Logger.log(`[${FUNC_NAME} ERROR] Gemini API FAILED for msg ${msgId}. Error: ${geminiApiResponse ? geminiApiResponse.error : 'Null or unexpected API response object'}`);
          writeErrorEntryToSheet_forLeads(leadsDataSheet, message, "Gemini API Call Fail (Leads)", geminiApiResponse ? String(geminiApiResponse.error).substring(0,500) : "Unknown API response error", leadsHeaderMap);
          allNewMessagesInThisThreadProcessedSuccessfully = false; // This message in thread had an API error
        }
      } catch (e) { // Catch script errors during processing of this message
        Logger.log(`[${FUNC_NAME} SCRIPT ERROR] Exception processing Msg ${msgId}: ${e.toString()}\nStack: ${e.stack}`);
        writeErrorEntryToSheet_forLeads(leadsDataSheet, message, "Script Error (Leads)", String(e.toString()).substring(0,500), leadsHeaderMap);
        allNewMessagesInThisThreadProcessedSuccessfully = false; // This message in thread had a script error
      }

      if (currentMessageHandledNoErrors) {
        processedLeadEmailIds.add(msgId); // Add to our run-time set to avoid re-processing in this same execution.
                                         // It will be written to sheet only if valid job data was found or if error was logged.
      }
      Utilities.sleep(1500 + Math.floor(Math.random() * 1000)); // Be respectful to APIs
    } // End loop for messages in a thread

    // Thread Relabeling Logic
    if (threadContainedAtLeastOneNewMessage) {
      if (allNewMessagesInThisThreadProcessedSuccessfully) {
        if (doneProcessLabel) { 
          try { thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel); Logger.log(`[${FUNC_NAME} INFO] Thread ${thread.getId()} successfully processed & moved to "${doneProcessLabelName}".`); }
          catch (eRelabel) { Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} relabel (success case) error: ${eRelabel.message}`); }
        } else { 
          try { thread.removeLabel(needsProcessLabel); Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} processed. Removed from "${needsProcessLabelName}", but "Done" label object is missing.`); }
          catch (eRemOnly) { Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} removeLabel error: ${eRemOnly.message}`); }
        }
      } else {
        Logger.log(`[${FUNC_NAME} WARN] Thread ${thread.getId()} contained new messages but encountered errors during processing. NOT moved from "${needsProcessLabelName}". Will be re-attempted.`);
      }
    } else if (messagesInThread.length > 0) { // Thread had messages, but all were already processed (found in processedLeadEmailIds set)
        Logger.log(`[${FUNC_NAME} INFO] Thread ${thread.getId()} contained only previously processed messages. Ensuring it's labeled correctly.`);
        if (doneProcessLabel && !thread.getLabels().map(l=>l.getName()).includes(doneProcessLabelName)) {
            try { thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel); } catch (eOldDone) {}
        } else if (!doneProcessLabel) {
            try { thread.removeLabel(needsProcessLabel); } catch (eOldRem) {}
        }
    } else if (messagesInThread.length === 0) { // Empty thread
        Logger.log(`[${FUNC_NAME} INFO] Thread ${thread.getId()} was empty. Removing from "${needsProcessLabelName}".`);
        try { thread.removeLabel(needsProcessLabel); } catch(eEmptyThread) {/*Minor*/}
    }
    Utilities.sleep(500); // Pause between threads
  } // End loop for threads
  
  Logger.log(`\n==== ${FUNC_NAME}: FINISHED (${new Date().toLocaleString()}) === Messages Attempted This Run: ${messagesProcessedThisRunCounter}. Total Time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}
