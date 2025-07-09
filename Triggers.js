// File: Triggers.gs
// Description: Contains functions for creating, verifying, and managing
// time-driven triggers for the project.

/**
 * Creates or verifies a time-based trigger for a given function to run every X hours.
 * @param {string} functionName The name of the function to trigger.
 * @param {number} hours The interval in hours.
 * @return {boolean} True if a new trigger was created, false if it already existed or an error occurred.
 */
function createTimeDrivenTrigger(functionName = 'processJobApplicationEmails', hours = 1) {
  let exists = false;
  let newTriggerCreated = false;
  try {
    const existingTriggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existingTriggers.length; i++) {
      if (existingTriggers[i].getHandlerFunction() === functionName && 
          existingTriggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
        // Could add more checks here, e.g., if the schedule matches `everyHours(hours)`
        // For now, simple existence check is used.
        exists = true;
        break;
      }
    }

    if (!exists) {
      ScriptApp.newTrigger(functionName).timeBased().everyHours(hours).create();
      Logger.log(`[INFO] TRIGGER: ${hours}-hourly trigger for "${functionName}" CREATED successfully.`);
      newTriggerCreated = true;
    } else {
      Logger.log(`[INFO] TRIGGER: ${hours}-hourly trigger for "${functionName}" ALREADY EXISTS.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] TRIGGER: Failed to create or verify ${hours}-hourly trigger for "${functionName}": ${e.message} (Stack: ${e.stack})`);
    return false; // Indicate error or no new creation
  }
  return newTriggerCreated;
}

/**
 * Creates or verifies a daily time-based trigger for a given function to run at a specific hour.
 * @param {string} functionName The name of the function to trigger.
 * @param {number} hour The hour of the day (0-23) in the script's timezone.
 * @return {boolean} True if a new trigger was created, false if it already existed or an error occurred.
 */
function createOrVerifyStaleRejectTrigger(functionName = 'markStaleApplicationsAsRejected', hour = 2) { // Default to 2 AM
  let exists = false;
  let newTriggerCreated = false;
  try {
    const existingTriggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < existingTriggers.length; i++) {
      if (existingTriggers[i].getHandlerFunction() === functionName &&
          existingTriggers[i].getEventType() === ScriptApp.EventType.CLOCK) {
        // Could add more detailed check for daily at specific hour
        exists = true;
        break;
      }
    }

    if (!exists) {
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyDays(1)
        .atHour(hour)
        .inTimezone(Session.getScriptTimeZone()) // Best practice
        .create();
      Logger.log(`[INFO] TRIGGER: Daily trigger for "${functionName}" (around ${hour}:00 script timezone) CREATED successfully.`);
      newTriggerCreated = true;
    } else {
      Logger.log(`[INFO] TRIGGER: Daily trigger for "${functionName}" ALREADY EXISTS.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] TRIGGER: Failed to create or verify daily trigger for "${functionName}": ${e.message} (Stack: ${e.stack})`);
    return false; // Indicate error or no new creation
  }
  return newTriggerCreated;
}
