// File: GeminiService.gs
// Description: Handles all interactions with the Google Gemini API for
// AI-powered parsing of email content to extract job application details and job leads.

// --- GEMINI API PARSING LOGIC ---
function callGemini_forApplicationDetails(emailSubject, emailBody, apiKey) {
  if (!apiKey) {
    Logger.log("[INFO] GEMINI_PARSE_APP: API Key not provided. Skipping Gemini call.");
    return null;
  }
  if ((!emailSubject || emailSubject.trim() === "") && (!emailBody || emailBody.trim() === "")) {
    Logger.log("[WARN] GEMINI_PARSE_APP: Both email subject and body are empty. Skipping Gemini call.");
    return null;
  }

  const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
  if (DEBUG_MODE) Logger.log(`[DEBUG] GEMINI_PARSE_APP: Using API Endpoint: ${API_ENDPOINT.split('key=')[0] + "key=..."}`);

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : ""; // Max 12k chars for body snippet

  // Constants from Config.gs are used here
  const prompt = `You are a highly specialized AI assistant expert in parsing job application-related emails for a tracking system. Your sole purpose is to analyze the provided email Subject and Body, and extract three key pieces of information: "company_name", "job_title", and "status". You MUST return this information ONLY as a single, valid JSON object, with no surrounding text, explanations, apologies, or markdown.

CRITICAL INSTRUCTIONS - READ AND FOLLOW CAREFULLY:

**PRIORITY 1: Determine Relevance - IS THIS A JOB APPLICATION UPDATE FOR THE RECIPIENT?**
- Your FIRST task is to assess if the email DIRECTLY relates to a job application previously submitted by the recipient, or an update to such an application.
- **IF THE EMAIL IS NOT APPLICATION-RELATED:** This includes general newsletters, marketing or promotional emails, sales pitches, webinar invitations, event announcements, account security alerts, password resets, bills/invoices, platform notifications not tied to a specific submitted application (e.g., "new jobs you might like"), or spam.
    - In such cases, IMMEDIATELY set ALL three fields ("company_name", "job_title", "status") to the exact string "${MANUAL_REVIEW_NEEDED}".
    - Do NOT attempt to extract any information from these irrelevant emails.
    - Your output for these MUST be: {"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${MANUAL_REVIEW_NEEDED}"}

**PRIORITY 2: If Application-Related, Proceed with Extraction:**

1.  "company_name":
    *   **Goal**: Extract the full, official name of the HIRING COMPANY to which the user applied.
    *   **ATS Handling**: Emails often originate from Applicant Tracking Systems (ATS) like Greenhouse (notifications@greenhouse.io), Lever (no-reply@hire.lever.co), Workday, Taleo, iCIMS, Ashby, SmartRecruiters, etc. The sender domain may be the ATS. You MUST identify the actual hiring company mentioned WITHIN the email subject or body. Look for phrases like "Your application to [Hiring Company]", "Careers at [Hiring Company]", "Update from [Hiring Company]", or the company name near the job title.
    *   **Do NOT extract**: The name of the ATS (e.g., "Greenhouse", "Lever"), the name of the job board (e.g., "LinkedIn", "Indeed", "Wellfound" - unless the job board IS the direct hiring company), or generic terms.
    *   **Ambiguity**: If the hiring company name is genuinely unclear from an application context, or only an ATS name is present without the actual company, use "${MANUAL_REVIEW_NEEDED}".
    *   **Accuracy**: Prefer full legal names if available (e.g., "Acme Corporation" over "Acme").

2.  "job_title":
    *   **Goal**: Extract the SPECIFIC job title THE USER APPLIED FOR, as mentioned in THIS email. The title is often explicitly stated after phrases like "your application for...", "application to the position of...", "the ... role", or directly alongside the company name in application submission/viewed confirmations.
    *   **LinkedIn Emails ("Application Sent To..." / "Application Viewed By...")**: These emails (often from sender "LinkedIn") frequently state the company name AND the job title the user applied for directly in the main body or a prominent header within the email content. Scrutinize these carefully for both. Example: "Your application for **Senior Product Manager** was sent to **Innovate Corp**." or "A recruiter from **Innovate Corp** viewed your application for **Senior Product Manager**." Extract "Senior Product Manager".
    *   **ATS Confirmation Emails (e.g., from Greenhouse, Lever)**: These emails confirming receipt of an application (e.g., "We've received your application to [Company]") often DO NOT restate the specific job title within the body of *that specific confirmation email*. If the job title IS NOT restated, you MUST use "${MANUAL_REVIEW_NEEDED}" for the job_title. Do not assume it from the subject line unless the subject clearly states "Your application for [Job Title] at [Company]".
    *   **General Updates/Rejections**: Some updates or rejections may or may not restate the title. If the title of the specific application is not clearly present in THIS email, use "${MANUAL_REVIEW_NEEDED}".
    *   **Strict Rule**: Do NOT infer a job title from company career pages, other listed jobs, or generic phrases like "various roles" unless that phrase directly follows "your application for". Only extract what is stated for THIS specific application event in THIS email. If in doubt, or if only a very generic descriptor like "a role" is used without specifics, prefer "${MANUAL_REVIEW_NEEDED}".

3.  "status":
    *   **Goal**: Determine the current status of the application based on the content of THIS email.
    *   **Strictly Adhere to List**: You MUST choose a status ONLY from the following exact list. Do not invent new statuses or use variations:
        *   "${DEFAULT_STATUS}" (Maps to: Application submitted, application sent, successfully applied, application received - first confirmation)
        *   "${REJECTED_STATUS}" (Maps to: Not moving forward, unfortunately, decided not to proceed, position filled by other candidates, regret to inform)
        *   "${OFFER_STATUS}" (Maps to: Offer of employment, pleased to offer, job offer)
        *   "${INTERVIEW_STATUS}" (Maps to: Invitation to interview, schedule an interview, interview request, like to speak with you)
        *   "${ASSESSMENT_STATUS}" (Maps to: Online assessment, coding challenge, technical test, skills test, take-home assignment)
        *   "${APPLICATION_VIEWED_STATUS}" (Maps to: Application was viewed by recruiter/company, your profile was viewed for the role)
        *   "Update/Other" (Maps to: General updates like "still reviewing applications," "we're delayed," "thanks for your patience," status is mentioned but unclear which of the above it fits best.)
    *   **Exclusion**: "${ACCEPTED_STATUS}" is typically set manually by the user after they accept an offer; do not use it.
    *   **Last Resort**: If the email is clearly job-application-related for the recipient, but the status is absolutely ambiguous and doesn't fit "Update/Other" (very rare), then as a final fallback, use "${MANUAL_REVIEW_NEEDED}" for the status.

**Output Requirements**:
*   **ONLY JSON**: Your entire response must be a single, valid JSON object.
*   **NO Extra Text**: No explanations, greetings, apologies, summaries, or markdown formatting (like \`\`\`json\`\`\`).
*   **Structure**: {"company_name": "...", "job_title": "...", "status": "..."}
*   **Placeholder Usage**: Adhere strictly to using "${MANUAL_REVIEW_NEEDED}" when information is absent or criteria are not met, as instructed for each field.

--- EXAMPLES START ---
Example 1 (LinkedIn "Application Sent To Company - Title Clearly Stated"):
Subject: Francis, your application was sent to MycoWorks
Body: LinkedIn. Your application was sent to MycoWorks. MycoWorks - Emeryville, CA (On-Site). Data Architect/Analyst. Applied on May 16, 2025.
Output:
{"company_name": "MycoWorks","job_title": "Data Architect/Analyst","status": "${DEFAULT_STATUS}"}

Example 2 (Indeed "Application Submitted", title present):
Subject: Indeed Application: Senior Software Engineer
Body: indeed. Application submitted. Senior Software Engineer. Innovatech Solutions - Anytown, USA. The following items were sent to Innovatech Solutions.
Output:
{"company_name": "Innovatech Solutions","job_title": "Senior Software Engineer","status": "${DEFAULT_STATUS}"}

Example 3 (Rejection from ATS, title might be in subject, but not confirmed in this email body):
Subject: Update on your application for Product Manager at MegaEnterprises
Body: From: no-reply@greenhouse.io. Dear Applicant, Thank you for your interest in MegaEnterprises. After careful consideration, we have decided to move forward with other candidates for this position.
Output:
{"company_name": "MegaEnterprises","job_title": "Product Manager","status": "${REJECTED_STATUS}"} 
(Self-correction: Title "Product Manager" taken from subject if directly linked to "your application". If subject was generic like "Application Update", job_title would be ${MANUAL_REVIEW_NEEDED})

Example 4 (Interview Invitation via ATS, title present):
Subject: Invitation to Interview: Data Analyst at Beta Innovations (via Lever)
Body: We were impressed with your application for the Data Analyst role and would like to invite you to an interview...
Output:
{"company_name": "Beta Innovations","job_title": "Data Analyst","status": "${INTERVIEW_STATUS}"}

Example 5 (ATS Email - Application Received, NO specific title in THIS email body):
Subject: Thank you for applying to Handshake!
Body: no-reply@greenhouse.io. Hi Francis, Thank you for your interest in Handshake! We have received your application and will be reviewing your background shortly... Handshake Recruiting.
Output:
{"company_name": "Handshake","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${DEFAULT_STATUS}"}

Example 6 (Unrelated Marketing):
Subject: Join our webinar on Future Tech!
Body: Hi User, Don't miss out on our exclusive webinar...
Output:
{"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${MANUAL_REVIEW_NEEDED}"}

Example 7 (LinkedIn "Application Viewed By..." - Title Clearly Stated):
Subject: Your application was viewed by Gotham Technology Group
Body: LinkedIn. Great job getting noticed by the hiring team at Gotham Technology Group. Gotham Technology Group - New York, United States. Business Analyst/Product Manager. Applied on May 14.
Output:
{"company_name": "Gotham Technology Group","job_title": "Business Analyst/Product Manager","status": "${APPLICATION_VIEWED_STATUS}"}

Example 8 (Wellfound "Application Submitted" - Often has title):
Subject: Application to LILT successfully submitted
Body: wellfound. Your application to LILT for the position of Lead Product Manager has been submitted! View your application. LILT.
Output:
{"company_name": "LILT","job_title": "Lead Product Manager","status": "${DEFAULT_STATUS}"}

Example 9 (Email indicating general interest/no specific role or company clear):
Subject: An interesting opportunity
Body: Hi Francis, Your profile on LinkedIn matches an opening we have. Would you be open to a quick chat? Regards, Recruiter.
Output:
{"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "Update/Other"}
--- EXAMPLES END ---

--- START OF EMAIL TO PROCESS ---
Subject: ${emailSubject}
Body:
${bodySnippet}
--- END OF EMAIL TO PROCESS ---
Output JSON:
`; // End of prompt template literal

  const payload = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": { "temperature": 0.2, "maxOutputTokens": 512, "topP": 0.95, "topK": 40 },
    "safetySettings": [ 
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const options = {'method':'post', 'contentType':'application/json', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};

  if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE_APP: Calling API for subj: "${emailSubject.substring(0,100)}". Prompt len (approx): ${prompt.length}`);
  let response; let attempt = 0; const maxAttempts = 2;

  while(attempt < maxAttempts){
    attempt++;
    try {
      response = UrlFetchApp.fetch(API_ENDPOINT, options);
      const responseCode = response.getResponseCode(); const responseBody = response.getContentText();
      if(DEBUG_MODE) Logger.log(`[DEBUG] GEMINI_PARSE_APP (Attempt ${attempt}): RC: ${responseCode}. Body(start): ${responseBody.substring(0,200)}`);

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
          let extractedJsonString = jsonResponse.candidates[0].content.parts[0].text.trim();
          if (extractedJsonString.startsWith("```json")) extractedJsonString = extractedJsonString.substring(7).trim();
          if (extractedJsonString.startsWith("```")) extractedJsonString = extractedJsonString.substring(3).trim();
          if (extractedJsonString.endsWith("```")) extractedJsonString = extractedJsonString.substring(0, extractedJsonString.length - 3).trim();
          
          if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE_APP: Cleaned JSON from API: ${extractedJsonString}`);
          try {
            const extractedData = JSON.parse(extractedJsonString);
            if (typeof extractedData.company_name !== 'undefined' && 
                typeof extractedData.job_title !== 'undefined' && 
                typeof extractedData.status !== 'undefined') {
              Logger.log(`[INFO] GEMINI_PARSE_APP: Success. C:"${extractedData.company_name}", T:"${extractedData.job_title}", S:"${extractedData.status}"`);
              return {
                  company: extractedData.company_name || MANUAL_REVIEW_NEEDED, 
                  title: extractedData.job_title || MANUAL_REVIEW_NEEDED, 
                  status: extractedData.status || MANUAL_REVIEW_NEEDED
              };
            } else {
              Logger.log(`[WARN] GEMINI_PARSE_APP: JSON from Gemini missing fields. Output: ${extractedJsonString}`);
              return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED};
            }
          } catch (e) {
            Logger.log(`[ERROR] GEMINI_PARSE_APP: Error parsing JSON: ${e.toString()}\nString: >>>${extractedJsonString}<<<`);
            return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED};
          }
        } else if (jsonResponse.promptFeedback?.blockReason) {
          Logger.log(`[ERROR] GEMINI_PARSE_APP: Prompt blocked. Reason: ${jsonResponse.promptFeedback.blockReason}. Details: ${JSON.stringify(jsonResponse.promptFeedback.safetyRatings)}`);
          return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:`Blocked: ${jsonResponse.promptFeedback.blockReason}`};
        } else {
          Logger.log(`[ERROR] GEMINI_PARSE_APP: API response structure unexpected. Body (start): ${responseBody.substring(0,500)}`);
          return null; 
        }
      } else if (responseCode === 429) {
        Logger.log(`[WARN] GEMINI_PARSE_APP: Rate limit (429). Attempt ${attempt}/${maxAttempts}. Waiting...`);
        if (attempt < maxAttempts) { Utilities.sleep(5000 + Math.floor(Math.random() * 5000)); continue; }
        else { Logger.log(`[ERROR] GEMINI_PARSE_APP: Max retries for rate limit.`); return null; }
      } else {
        Logger.log(`[ERROR] GEMINI_PARSE_APP: API HTTP error. Code: ${responseCode}. Body (start): ${responseBody.substring(0,500)}`);
        if (responseCode === 404 && responseBody.includes("is not found for API version")) {
            Logger.log(`[FATAL] GEMINI_MODEL_ERROR_APP: Model ${API_ENDPOINT.split('/models/')[1].split(':')[0]} not found.`)
        }
        return null;
      }
    } catch (e) {
      Logger.log(`[ERROR] GEMINI_PARSE_APP: Exception during API call (Attempt ${attempt}): ${e.toString()}\nStack: ${e.stack}`);
      if (attempt < maxAttempts) { Utilities.sleep(3000); continue; }
      return null;
    }
  }
  Logger.log(`[ERROR] GEMINI_PARSE_APP: Failed after ${maxAttempts} attempts.`);
  return null;
}


function callGemini_forJobLeads(emailBody, apiKey) {
    if (typeof emailBody !== 'string') {
        Logger.log(`[GEMINI_LEADS CRITICAL ERR] emailBody not string. Type: ${typeof emailBody}`);
        return { success: false, data: null, error: `emailBody is not a string.` };
    }

    const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;

    // Mock data logic for when API key is placeholder or not set (as before)
    if (!apiKey || apiKey === 'AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' || apiKey.trim() === '') {
        Logger.log("[GEMINI_LEADS WARN STUB] API Key placeholder/missing. Using MOCK for job leads.");
        // ... (your existing mock response logic for leads) ...
        if (emailBody.toLowerCase().includes("multiple job listings inside") || emailBody.toLowerCase().includes("software engineer at google")) {
            return { success: true, data: { candidates: [{ content: { parts: [{ text: JSON.stringify([ { "jobTitle": "Software Engineer (Mock)", "company": "Tech Alpha (Mock)", "location": "Remote", "source": "Mock Job Board", "jobUrl": "https://example.com/job/alpha", "notes": "This is a mock note." }, { "jobTitle": "Product Manager (Mock)", "company": "Innovate Beta (Mock)", "location": "New York, NY", "source": "Mock Alerts", "jobUrl": "https://example.com/job/beta", "notes": "Requires 5 years experience." } ]) }] } }] }, error: null };
        }
        return { success: true, data: { candidates: [{ content: { parts: [{ text: JSON.stringify([{ "jobTitle": "N/A (Mock Single)", "company": "Some Corp (Mock)", "location": "Remote", "source": "Mock Direct", "jobUrl": "N/A", "notes": "Basic mock entry." }]) }] } }] }, error: null };
    }

    // --- MODIFIED PROMPT ---
    const promptText = `You are an expert AI assistant specializing in extracting job posting details from email content, typically from job alerts or direct emails containing job opportunities.
From the following "Email Content", identify each distinct job posting.

For EACH job posting found, extract the following details:
- "jobTitle": The specific title of the job role (e.g., "Senior Software Engineer", "Product Marketing Manager").
- "company": The name of the hiring company.
- "location": The primary location of the job (e.g., "San Francisco, CA", "Remote", "London, UK", "Hybrid - New York").
- "source": If identifiable from the email content, the origin or job board where this posting was listed (e.g., "LinkedIn Job Alert", "Indeed", "Wellfound", "Company Careers Page" if mentioned). If not explicitly stated, use "N/A".
- "jobUrl": A direct URL link to the job application page or a more detailed job description, if present in the email. If no direct link for *this specific job* is found, use "N/A".
- "notes": Briefly extract 2-3 key requirements, responsibilities, or unique aspects mentioned for this specific job if readily available in the email text (e.g., "Requires Python & AWS; 5+ yrs exp", "Focus on B2B SaaS marketing", "Fast-paced startup environment"). Keep notes concise (max 150 characters). If no specific details are easily extractable for this job, use "N/A".

Strict Formatting Instructions:
- Your entire response MUST be a single, valid JSON array.
- Each element of the array MUST be a JSON object representing one job posting.
- Each JSON object MUST have exactly these keys: "jobTitle", "company", "location", "source", "jobUrl", "notes".
- If a specific field for a job is not found or not applicable, its value MUST be the string "N/A".
- If no job postings at all are found in the email content, return an empty JSON array: [].
- Do NOT include any text, explanations, apologies, or markdown (like \`\`\`json\`\`\`) before or after the JSON array.

--- EXAMPLE OUTPUT START (for an email with two jobs) ---
[
  {
    "jobTitle": "Senior Frontend Developer",
    "company": "Innovatech Solutions",
    "location": "Remote (US)",
    "source": "LinkedIn Job Alert",
    "jobUrl": "https://linkedin.com/jobs/view/12345",
    "notes": "React, TypeScript, Agile environment. 5+ years experience. UI/UX focus."
  },
  {
    "jobTitle": "Data Scientist",
    "company": "Alpha Analytics Co.",
    "location": "Boston, MA",
    "source": "Direct Email from Recruiter",
    "jobUrl": "N/A",
    "notes": "Machine learning, Python, SQL. PhD preferred. Early-stage startup."
  }
]
--- EXAMPLE OUTPUT END ---

Email Content:
---
${emailBody.substring(0, 30000)} 
---
JSON Array Output:`; // Max characters for body increased slightly

    const payload = {
        contents: [{ parts: [{ "text": promptText }] }],
        generationConfig: { 
            temperature: 0.2, 
            maxOutputTokens: 8192, // Kept high for potentially multiple listings
            // responseMimeType: "application/json" // Can try adding this if LLM still includes ```json
        },
        safetySettings: [ 
          { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
          { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
          { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
          { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
        ]
    };
    const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
    let attempt = 0; const maxAttempts = 2; // Or 3 for more resilience

    Logger.log(`[GEMINI_LEADS INFO] Calling Gemini for leads. Prompt length (approx): ${promptText.length}`);

    while (attempt < maxAttempts) {
        attempt++;
        try {
            const response = UrlFetchApp.fetch(API_ENDPOINT, options);
            const responseCode = response.getResponseCode(); 
            const responseBody = response.getContentText();
            Logger.log(`[GEMINI_LEADS DEBUG Attempt ${attempt}] RC: ${responseCode}. Body (start): ${responseBody.substring(0, 250)}...`);

            if (responseCode === 200) {
                try { 
                    // The response from Gemini might already be the JSON part.
                    // parseGeminiResponse_forJobLeads will handle cleaning ```json
                    return { success: true, data: JSON.parse(responseBody), error: null }; 
                }
                catch (jsonParseError) {
                    Logger.log(`[GEMINI_LEADS API ERROR] Parse Gemini JSON after 200 OK: ${jsonParseError}. Raw Body: ${responseBody}`);
                    return { success: false, data: null, error: `Parse API JSON (200 OK): ${jsonParseError}. Response: ${responseBody.substring(0,500)}` };
                }
            } else if (responseCode === 429 && attempt < maxAttempts) { // Rate limit
                Logger.log(`[GEMINI_LEADS API WARN ${responseCode}] Rate limit (attempt ${attempt}/${maxAttempts}). Waiting 3-5s...`);
                Utilities.sleep(3000 + Math.random() * 2000); 
                continue;
            } else { // Other HTTP errors
                Logger.log(`[GEMINI_LEADS API ERROR ${responseCode}] Full error: ${responseBody}`);
                const parsedError = JSON.parse(responseBody); // Try to parse error for details
                if (parsedError && parsedError.error && parsedError.error.message) {
                     return { success: false, data: null, error: `API Error ${responseCode}: ${parsedError.error.message}` };
                }
                return { success: false, data: null, error: `API Error ${responseCode}: ${responseBody.substring(0,500)}` };
            }
        } catch (e) { // Catch UrlFetchApp.fetch exceptions
            Logger.log(`[GEMINI_LEADS API CATCH (Attempt ${attempt}/${maxAttempts})] Fetch Error: ${e.toString()}`);
            if (attempt < maxAttempts) { Utilities.sleep(2000 + Math.random()*1000); continue; }
            return { success: false, data: null, error: `Fetch Error after ${maxAttempts} attempts: ${e.toString()}` };
        }
    }
    Logger.log(`[GEMINI_LEADS ERROR] Exceeded max retries for Gemini API.`);
    return { success: false, data: null, error: `Exceeded max retries (${maxAttempts}) for Gemini API.` };
}

// --- You will ALSO need to update `parseGeminiResponse_forJobLeads` ---
function parseGeminiResponse_forJobLeads(apiResponseData) {
    let jobListings = [];
    const FUNC_NAME = "parseGeminiResponse_forJobLeads";
    try {
        let jsonStringFromLLM = "";
        if (apiResponseData?.candidates?.[0]?.content?.parts?.[0]?.text) {
            jsonStringFromLLM = apiResponseData.candidates[0].content.parts[0].text.trim();
            // Clean potential markdown fences
            if (jsonStringFromLLM.startsWith("```json")) jsonStringFromLLM = jsonStringFromLLM.substring(7, jsonStringFromLLM.endsWith("```") ? jsonStringFromLLM.length - 3 : undefined).trim();
            else if (jsonStringFromLLM.startsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(3, jsonStringFromLLM.endsWith("```") ? jsonStringFromLLM.length - 3 : undefined).trim();
            else if (jsonStringFromLLM.endsWith("```")) jsonStringFromLLM = jsonStringFromLLM.substring(0, jsonStringFromLLM.length - 3).trim();
        } else {
            Logger.log(`[${FUNC_NAME} WARN] No parsable content string in Gemini response for leads.`);
            if (apiResponseData?.promptFeedback?.blockReason) {
                Logger.log(`[${FUNC_NAME} WARN] Prompt Block Reason: ${apiResponseData.promptFeedback.blockReason}.`);
            }
            return jobListings; // Empty array
        }

        Logger.log(`[${FUNC_NAME} DEBUG] Cleaned JSON string from LLM: ${jsonStringFromLLM.substring(0, 500)}...`);
        try {
            const parsedData = JSON.parse(jsonStringFromLLM);
            if (Array.isArray(parsedData)) {
                parsedData.forEach(job => {
                    if (job && typeof job === 'object' && (job.jobTitle || job.company)) { // Basic validation for a job object
                        jobListings.push({
                            jobTitle: job.jobTitle || "N/A", 
                            company: job.company || "N/A",
                            location: job.location || "N/A", 
                            source: job.source || "N/A",         // <<< NEW
                            jobUrl: job.jobUrl || "N/A",         // Changed from linkToJobPosting
                            notes: job.notes || "N/A"            // <<< NEW
                        });
                    } else { Logger.log(`[${FUNC_NAME} WARN] Skipped invalid item in parsed job listings array: ${JSON.stringify(job)}`); }
                });
            } else if (typeof parsedData === 'object' && parsedData !== null && (parsedData.jobTitle || parsedData.company)) { // Handle case where LLM returns a single object instead of array
                jobListings.push({
                    jobTitle: parsedData.jobTitle || "N/A", 
                    company: parsedData.company || "N/A",
                    location: parsedData.location || "N/A", 
                    source: parsedData.source || "N/A",      // <<< NEW
                    jobUrl: parsedData.jobUrl || "N/A",      // Changed from linkToJobPosting
                    notes: parsedData.notes || "N/A"         // <<< NEW
                });
                Logger.log(`[${FUNC_NAME} WARN] LLM output was a single object, parsed as one job.`);
            } else { 
                Logger.log(`[${FUNC_NAME} WARN] LLM output not a JSON array or a single parsable job object. Output (start): ${jsonStringFromLLM.substring(0, 200)}`); 
            }
        } catch (jsonError) {
            Logger.log(`[${FUNC_NAME} ERROR] Failed to parse JSON string from LLM: ${jsonError}. String (start): ${jsonStringFromLLM.substring(0, 500)}`);
        }
        Logger.log(`[${FUNC_NAME} INFO] Successfully parsed ${jobListings.length} job listings from Gemini response.`);
        return jobListings;
    } catch (e) {
        Logger.log(`[${FUNC_NAME} ERROR] Outer error during parsing Gemini response for leads: ${e.toString()}. API Resp Data (partial): ${JSON.stringify(apiResponseData).substring(0, 300)}`);
        return jobListings; // Return empty array on error
    }
}
