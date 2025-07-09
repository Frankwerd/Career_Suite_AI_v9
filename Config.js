// File: Config.gs
// Project: CareerSuite.AI Job Tracker
// Description: Contains all global configuration constants for the project.

// --- General Configuration ---
const DEBUG_MODE = true; // Set to false for production to reduce logging

// --- BRANDING COLORS ---
const BRAND_COLORS = {
  LAPIS_LAZULI: "#33658A",    // Primary (e.g., "Applications" tab, Dashboard Title BG)
  CAROLINA_BLUE: "#86BBD8",   // Secondary (e.g., "Dashboard" tab, Banding Header)
  CHARCOAL: "#2F4858",        // Text, Dark Backgrounds, "HelperData" tab
  HUNYADI_YELLOW: "#F6AE2D",  // Accent 1 (e.g., "Potential Job Leads" tab & its Banding Header)
  PALE_ORANGE: "#F26419",      // Accent 2 (e.g., for some UI elements or specific highlights)
  WHITE: "#FFFFFF",           // Text color on dark backgrounds, First Banding Row
  PALE_GREY: "#F0F4F8",       // Page/Card Light Background (for dashboard cells), Second Banding Row
  MEDIUM_GREY_BORDER: "#DDE4E9" // Borders (for dashboard cells)
};

const CUSTOM_MENU_NAME = "⚙️ CareerSuite.AI Tools";

// --- Gemini AI Configuration ---
const GEMINI_API_KEY_PROPERTY = 'CAREERSUITE_GEMINI_API_KEY';

// --- Spreadsheet Configuration ---
const TARGET_SPREADSHEET_FILENAME = "CareerSuite.ai Data";
const FIXED_SPREADSHEET_ID = ""; // For manual runs if specific sheet targeted

// --- Template Sheet ID (Used by WebApp to create a new user sheet) ---
const TEMPLATE_SHEET_ID = "12jj5lTyu_MzA6KBkfD-30mj-KYHaX-BjouFMtPIIzFc"; // YOUR ACTUAL TEMPLATE ID

// --- Tab Names within the Main Spreadsheet ---
const APP_TRACKER_SHEET_TAB_NAME = "Applications";
const DASHBOARD_TAB_NAME = "Dashboard";
const HELPER_SHEET_NAME = "DashboardHelperData";
const LEADS_SHEET_TAB_NAME = "Potential Job Leads";

// --- Gmail Label Configuration (CareerSuite.AI Branded) ---
const MASTER_GMAIL_LABEL_PARENT = "CareerSuite.AI";
const TRACKER_GMAIL_LABEL_PARENT = MASTER_GMAIL_LABEL_PARENT + "/Applications";
const TRACKER_GMAIL_LABEL_TO_PROCESS = TRACKER_GMAIL_LABEL_PARENT + "/To Process";
const TRACKER_GMAIL_LABEL_PROCESSED = TRACKER_GMAIL_LABEL_PARENT + "/Processed";
const TRACKER_GMAIL_LABEL_MANUAL_REVIEW = TRACKER_GMAIL_LABEL_PARENT + "/Manual Review";
const TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES = 'subject:("your application" OR "application to" OR "application for" OR "application update" OR "thank you for applying" OR "thanks for applying" OR "thank you for your interest" OR "received your application")';
const LEADS_GMAIL_LABEL_PARENT = MASTER_GMAIL_LABEL_PARENT + "/Leads";
const LEADS_GMAIL_LABEL_TO_PROCESS = LEADS_GMAIL_LABEL_PARENT + "/To Process";
const LEADS_GMAIL_LABEL_PROCESSED = LEADS_GMAIL_LABEL_PARENT + "/Processed";
const LEADS_GMAIL_FILTER_QUERY = 'subject:("job alert" OR "job opportunity" OR "new role matching your profile") OR from:(*@linkedin.com AND (subject:(jobs OR job) OR "matching your profile"))';

// --- "Applications" Sheet Columns & Headers ---
const PROCESSED_TIMESTAMP_COL = 1; const EMAIL_DATE_COL = 2; const PLATFORM_COL = 3; const COMPANY_COL = 4; const JOB_TITLE_COL = 5; const STATUS_COL = 6; const PEAK_STATUS_COL = 7; const LAST_UPDATE_DATE_COL = 8; const EMAIL_SUBJECT_COL = 9; const EMAIL_LINK_COL = 10; const EMAIL_ID_COL = 11; const NOTES_COL = 12;
const TOTAL_COLUMNS_IN_APP_SHEET = 12;
const APP_TRACKER_SHEET_HEADERS = ["Processed Timestamp", "Email Date", "Platform", "Company", "Job Title", "Status", "Peak Status", "Last Update Date", "Email Subject", "Email Link", "Email ID", "Notes"];
const APP_SHEET_COLUMN_WIDTHS = [{ col: PROCESSED_TIMESTAMP_COL, width: 160 }, { col: EMAIL_DATE_COL, width: 120 }, { col: PLATFORM_COL, width: 100 }, { col: COMPANY_COL, width: 200 }, { col: JOB_TITLE_COL, width: 250 }, { col: STATUS_COL, width: 150 }, { col: PEAK_STATUS_COL, width: 150 }, { col: LAST_UPDATE_DATE_COL, width: 160 }, { col: EMAIL_SUBJECT_COL, width: 300 }, { col: EMAIL_LINK_COL, width: 80 }, { col: EMAIL_ID_COL, width: 180 }, { col: NOTES_COL, width: 250 }];

// --- Status Values & Hierarchy ---
const DEFAULT_STATUS = "Applied"; const REJECTED_STATUS = "Rejected"; const OFFER_STATUS = "Offer Received"; const ACCEPTED_STATUS = "Offer Accepted"; const INTERVIEW_STATUS = "Interview Scheduled"; const ASSESSMENT_STATUS = "Assessment/Screening"; const APPLICATION_VIEWED_STATUS = "Application Viewed"; const MANUAL_REVIEW_NEEDED = "N/A - Manual Review"; const DEFAULT_PLATFORM = "Other";
const STATUS_HIERARCHY = { [MANUAL_REVIEW_NEEDED]: -1, "Update/Other": 0, [DEFAULT_STATUS]: 1, [APPLICATION_VIEWED_STATUS]: 2, [ASSESSMENT_STATUS]: 3, [INTERVIEW_STATUS]: 4, [OFFER_STATUS]: 5, [REJECTED_STATUS]: 5, [ACCEPTED_STATUS]: 6 };

// --- Auto-Reject Stale Applications Config ---
const WEEKS_THRESHOLD = 7; const FINAL_STATUSES_FOR_STALE_CHECK = new Set([REJECTED_STATUS, ACCEPTED_STATUS, "Withdrawn"]);

// --- Email Parsing Keywords ---
const REJECTION_KEYWORDS = ["unfortunately", "regret to inform", "not moving forward", "decided not to proceed", "other candidates", "filled the position", "thank you for your time but"]; const OFFER_KEYWORDS = ["pleased to offer", "offer of employment", "job offer", "formally offer you the position"]; const INTERVIEW_KEYWORDS = ["invitation to interview", "schedule an interview", "interview request", "like to speak with you", "next steps involve an interview", "interview availability"]; const ASSESSMENT_KEYWORDS = ["assessment", "coding challenge", "online test", "technical screen", "next step is a skill assessment", "take a short test"]; const APPLICATION_VIEWED_KEYWORDS = ["application was viewed", "your application was viewed by", "recruiter viewed your application", "company viewed your application", "viewed your profile for the role"]; const PLATFORM_DOMAIN_KEYWORDS = { "linkedin.com": "LinkedIn", "indeed.com": "Indeed", "wellfound.com": "Wellfound", "angel.co": "Wellfound", "otta.com": "Otta" }; const IGNORED_DOMAINS = new Set(['greenhouse.io', 'lever.co', 'myworkday.com', 'icims.com', 'ashbyhq.com', 'smartrecruiters.com', 'bamboohr.com', 'taleo.net', 'gmail.com', 'google.com', 'example.com']);

// --- "DashboardHelperData" Sheet Headers & Column Widths ---
const DASHBOARD_HELPER_HEADERS = ["Category", "Metric/Platform", "Count", "Category2", "Metric/Week", "Count2", "Category3", "Metric/Stage", "Count3"]; // Example structure
const HELPER_SHEET_COLUMN_WIDTHS = [{ col: 1, width: 120 }, { col: 2, width: 180 }, { col: 3, width: 80 }, { col: 4, width: 20 }/*Spacer*/, { col: 5, width: 120 }, { col: 6, width: 180 }, { col: 7, width: 80 }, { col: 8, width: 20 }/*Spacer*/, { col: 9, width: 120 }, { col: 10, width: 180 }, { col: 11, width: 80 }];

// --- "Potential Job Leads" Sheet Headers & Column Widths ---
const LEADS_SHEET_HEADERS = [
    "Date Added", "Job Title", "Company", "Location", 
    "Source", // For Gemini extracted source like "LinkedIn Job Alert"
    "Job URL", 
    "Status", 
    "Notes", // For Gemini extracted notes
    "Applied Date", // Typically user input
    "Follow-up Date", // Typically user input
    "Source Email ID", // Script input
    "Processed Timestamp", // Script input
    // "Source Email Subject" // Optional: if you want subject and Gemini's extracted source separately
];
const TOTAL_COLUMNS_IN_LEADS_SHEET = LEADS_SHEET_HEADERS.length;
// And update LEADS_SHEET_COLUMN_WIDTHS accordingly
const LEADS_SHEET_COLUMN_WIDTHS = [{ col: 1, width: 100 } /*Date Added*/, { col: 2, width: 220 } /*Job Title*/, { col: 3, width: 180 } /*Company*/, { col: 4, width: 150 } /*Location*/, { col: 5, width: 180 } /*Source*/, { col: 6, width: 250 } /*Job URL*/, { col: 7, width: 100 } /*Status*/, { col: 8, width: 250 } /*Notes*/, { col: 9, width: 100 }/*Applied Date*/, { col: 10, width: 100 }/*Follow-up Date*/, { col: 11, width: 150 }/*Source Email ID*/, { col: 12, width: 160 }/*Processed Timestamp*/];

// --- Job Leads Tracker: UserProperty KEYS ---
const LEADS_USER_PROPERTY_TO_PROCESS_LABEL_ID = 'leadsGmailToProcessLabelId_vCSAI';
const LEADS_USER_PROPERTY_PROCESSED_LABEL_ID = 'leadsGmailProcessedLabelId_vCSAI';
