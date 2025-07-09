// File: Dashboard.js
// Description: The final, complete version of the Dashboard script, incorporating all
// formatting, layout, and formula corrections based on the user's golden standard.

// --- Helper Functions ---
function BRAND_COLORS_CHART_ARRAY() {
    return [
        BRAND_COLORS.LAPIS_LAZULI, BRAND_COLORS.CAROLINA_BLUE, BRAND_COLORS.HUNYADI_YELLOW,
        BRAND_COLORS.PALE_ORANGE, BRAND_COLORS.CHARCOAL, "#27AE60", "#8E44AD"
    ];
}

function _columnToLetter_DashboardLocal(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// --- Sheet Creation ---
function getOrCreateDashboardSheet(spreadsheet) {
    let sheet = spreadsheet.getSheetByName(DASHBOARD_TAB_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(DASHBOARD_TAB_NAME);
    }
    sheet.setTabColor(BRAND_COLORS.CAROLINA_BLUE);
    return sheet;
}

function getOrCreateHelperSheet(spreadsheet) {
    let sheet = spreadsheet.getSheetByName(HELPER_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(HELPER_SHEET_NAME);
    }
    // Hiding and coloring the helper sheet is handled in the main setup flow.
    return sheet;
}

// --- Formula and Formatting Setup ---

/**
 * Sets up the formulas in the DashboardHelperData sheet, using the original multi-step logic.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} helperSheet The "DashboardHelperData" sheet object.
 */
function setupHelperSheetFormulas(helperSheet) {
  const FUNC_NAME = "setupHelperSheetFormulas_OriginalLogic";
  if (!helperSheet) { return false; }
  Logger.log(`[${FUNC_NAME}] Setting up original formulas in "${helperSheet.getName()}".`);

  try {
    helperSheet.clearContents();

    const appSheetNameForFormula = `'${APP_TRACKER_SHEET_TAB_NAME}'!`;
    const platformColLetter = _columnToLetter_DashboardLocal(PLATFORM_COL);
    const emailDateColLetter = _columnToLetter_DashboardLocal(EMAIL_DATE_COL);
    const peakStatusColLetter = _columnToLetter_DashboardLocal(PEAK_STATUS_COL);
    const companyColLetter = _columnToLetter_DashboardLocal(COMPANY_COL);

    // Platform Distribution Data
    helperSheet.getRange("A1").setValue("Platform");
    helperSheet.getRange("B1").setValue("Count");
    helperSheet.getRange("A2").setFormula(`=IFERROR(QUERY(${appSheetNameForFormula}${platformColLetter}2:${platformColLetter}, "SELECT Col1, COUNT(Col1) WHERE Col1 IS NOT NULL AND Col1 <> '' GROUP BY Col1 ORDER BY COUNT(Col1) DESC LABEL Col1 '', COUNT(Col1) ''", 0), {"No Data",0})`);

    // Applications Over Time (Weekly) - Using Original Multi-Step Logic
    helperSheet.getRange("J1").setValue("RAW_VALID_DATES_FOR_WEEKLY");
    helperSheet.getRange("J2").setFormula(`=IFERROR(FILTER(${appSheetNameForFormula}${emailDateColLetter}2:${emailDateColLetter}, ISNUMBER(${appSheetNameForFormula}${emailDateColLetter}2:${emailDateColLetter})), "")`);
    helperSheet.getRange("K1").setValue("CALCULATED_WEEK_STARTS (Mon)");
    helperSheet.getRange("K2").setFormula(`=ARRAYFORMULA(IF(ISBLANK(J2:J), "", DATE(YEAR(J2:J), MONTH(J2:J), DAY(J2:J) - WEEKDAY(J2:J, 2) + 1)))`);
    
    helperSheet.getRange("D1").setValue("Week Starting");
    helperSheet.getRange("D2").setFormula(`=IFERROR(SORT(UNIQUE(FILTER(K2:K, K2:K<>""))), {"No Date Data"})`);
    helperSheet.getRange("D2:D").setNumberFormat("M/d/yyyy");
    helperSheet.getRange("E1").setValue("Applications");
    helperSheet.getRange("E2").setFormula(`=ARRAYFORMULA(IF(D2:D="", "", COUNTIF(K2:K, D2:D)))`);
    
    // Application Funnel
    helperSheet.getRange("G1").setValue("Stage"); 
    helperSheet.getRange("H1").setValue("Count");
    const funnelStagesValues = [DEFAULT_STATUS, APPLICATION_VIEWED_STATUS, ASSESSMENT_STATUS, INTERVIEW_STATUS, OFFER_STATUS];
    helperSheet.getRange(2, 7, funnelStagesValues.length, 1).setValues(funnelStagesValues.map(stage => [stage]));
    helperSheet.getRange("H2").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}${companyColLetter}2:${companyColLetter}),0)`); 
    for (let i = 1; i < funnelStagesValues.length; i++) {
        helperSheet.getRange(i + 2, 8).setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}${peakStatusColLetter}2:${peakStatusColLetter}, G${i + 2}),0)`);
    }
    
    SpreadsheetApp.flush();
    return true;
  } catch (e) {
    Logger.log(`[${FUNC_NAME} ERROR] ${e.toString()}`);
    return false;
  }
}

/**
 * Formats the dashboard sheet layout, scorecards, and formulas to precisely match the target design.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The dashboard sheet object.
 */
function formatDashboardSheet(dashboardSheet) {
    const FUNC_NAME = "formatDashboardSheet_GoldenStandard";
    if (!dashboardSheet) { return false; }

    try {
        dashboardSheet.clear();
        dashboardSheet.setHiddenGridlines(true);

        const MAIN_TITLE_BG = BRAND_COLORS.LAPIS_LAZULI;
        const HEADER_TEXT_COLOR = BRAND_COLORS.WHITE;
        const CARD_BG = BRAND_COLORS.PALE_GREY;
        const CARD_TEXT_COLOR = BRAND_COLORS.CHARCOAL;
        const CARD_BORDER_COLOR = BRAND_COLORS.MEDIUM_GREY_BORDER;
        const PRIMARY_VALUE_COLOR = BRAND_COLORS.PALE_ORANGE;
        const SECONDARY_VALUE_COLOR = BRAND_COLORS.HUNYADI_YELLOW;

        const spacerColAWidth = 20, labelWidth = 150, valueWidth = 75, spacerS = 15;

        dashboardSheet.getRange("A1:M1").merge().setValue("CareerSuite.AI Job Application Dashboard")
            .setBackground(MAIN_TITLE_BG).setFontColor(HEADER_TEXT_COLOR).setFontSize(18).setFontWeight("bold")
            .setHorizontalAlignment("center").setVerticalAlignment("middle");
        dashboardSheet.setRowHeight(1, 45);

        dashboardSheet.getRange("B3").setValue("Key Metrics Overview:").setFontSize(14).setFontWeight("bold");
        
        // Setting row heights to match the golden standard's spacing
        [2, 4, 6, 8, 10, 12, 28].forEach(r => dashboardSheet.setRowHeight(r, 10)); // Original code had 29 for funnel, let's use 28 for title
        [3, 11, 29].forEach(r => dashboardSheet.setRowHeight(r, 25)); // Original code had these as title rows
        [5, 7, 9].forEach(r => dashboardSheet.setRowHeight(r, 40));

        const appSheetName = `'${APP_TRACKER_SHEET_TAB_NAME}'`;
        const companyColLetter = _columnToLetter_DashboardLocal(COMPANY_COL);
        const statusColLetter = _columnToLetter_DashboardLocal(STATUS_COL);
        const peakStatusColLetter = _columnToLetter_DashboardLocal(PEAK_STATUS_COL);

        const manualReviewFormula = `=IFERROR(SUM(ARRAYFORMULA(SIGN((${appSheetName}!D2:D="${MANUAL_REVIEW_NEEDED}")+(${appSheetName}!E2:E="${MANUAL_REVIEW_NEEDED}")+(${appSheetName}!F2:F="${MANUAL_REVIEW_NEEDED}")))),0)`;
        
        const metrics = [
            { label: "Total Apps", valueFormula: `=IFERROR(COUNTA(${appSheetName}!${companyColLetter}2:${companyColLetter}), 0)`, labelCell: "B5", valueCell: "C5" },
            { label: "Peak Interviews", valueFormula: `=IFERROR(COUNTIF(${appSheetName}!${peakStatusColLetter}2:${peakStatusColLetter},"${INTERVIEW_STATUS}"), 0)`, labelCell: "E5", valueCell: "F5" },
            { label: "Interview Rate", valueFormula: `=IFERROR(F5/C5, 0)`, format: "0.00%", color: SECONDARY_VALUE_COLOR, labelCell: "H5", valueCell: "I5" },
            { label: "Offer Rate", valueFormula: `=IFERROR(F7/C5, 0)`, format: "0.00%", color: SECONDARY_VALUE_COLOR, labelCell: "K5", valueCell: "L5" },
            { label: "Active Apps", valueFormula: `=IFERROR(COUNTIFS(${appSheetName}!${statusColLetter}2:${statusColLetter}, "<>"&"", ${appSheetName}!${statusColLetter}2:${statusColLetter}, "<>${REJECTED_STATUS}", ${appSheetName}!${statusColLetter}2:${statusColLetter}, "<>${ACCEPTED_STATUS}"), 0)`, labelCell: "B7", valueCell: "C7" },
            { label: "Peak Offers", valueFormula: `=IFERROR(COUNTIF(${appSheetName}!${peakStatusColLetter}2:${peakStatusColLetter},"${OFFER_STATUS}"), 0)`, labelCell: "E7", valueCell: "F7" },
            { label: "Current Interviews", valueFormula: `=IFERROR(COUNTIF(${appSheetName}!${statusColLetter}2:${statusColLetter},"${INTERVIEW_STATUS}"), 0)`, labelCell: "H7", valueCell: "I7" },
            { label: "Current Assessments", valueFormula: `=IFERROR(COUNTIF(${appSheetName}!${statusColLetter}2:${statusColLetter},"${ASSESSMENT_STATUS}"), 0)`, labelCell: "K7", valueCell: "L7" },
            { label: "Total Rejections", valueFormula: `=IFERROR(COUNTIF(${appSheetName}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}"), 0)`, labelCell: "B9", valueCell: "C9" },
            { label: "Apps Viewed (Peak)", valueFormula: `=IFERROR(COUNTIF(${appSheetName}!${peakStatusColLetter}2:${peakStatusColLetter},"${APPLICATION_VIEWED_STATUS}"),0)`, labelCell: "E9", valueCell: "F9" },
            { label: "Manual Review", valueFormula: manualReviewFormula, color: SECONDARY_VALUE_COLOR, labelCell: "H9", valueCell: "I9" },
            { label: "Direct Reject Rate", valueFormula: `=IFERROR(COUNTIFS(${appSheetName}!${peakStatusColLetter}2:${peakStatusColLetter},"${DEFAULT_STATUS}",${appSheetName}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}")/C5, 0)`, format: "0.00%", color: SECONDARY_VALUE_COLOR, labelCell: "K9", valueCell: "L9" }
        ];

        metrics.forEach(metric => {
            dashboardSheet.getRange(metric.labelCell).setValue(metric.label).setFontWeight("bold").setVerticalAlignment("middle");
            dashboardSheet.getRange(metric.valueCell).setFormula(metric.valueFormula).setFontSize(15).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setFontColor(metric.color || PRIMARY_VALUE_COLOR).setNumberFormat(metric.format || "0");
            dashboardSheet.getRange(metric.labelCell + ":" + metric.valueCell).setBackground(CARD_BG).setBorder(true, true, true, true, true, true, CARD_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
        });

        dashboardSheet.getRange("B11").setValue("Platform & Weekly Trends").setFontSize(12).setFontWeight("bold");
        dashboardSheet.getRange("B28").setValue("Application Funnel Analysis").setFontSize(12).setFontWeight("bold");

        if (dashboardSheet.getMaxColumns() > 13) {
            dashboardSheet.hideColumns(14, dashboardSheet.getMaxColumns() - 13);
        }

        return true;
    } catch (e) {
        Logger.log(`[${FUNC_NAME} ERROR] ${e.toString()}`);
        return false;
    }
}


function updateDashboardMetrics(dashboardSheet, helperSheet) {
    const FUNC_NAME = "updateDashboardMetrics";
    Logger.log(`[${FUNC_NAME}] Rebuilding all dashboard charts...`);
    
    dashboardSheet.getCharts().forEach(chart => dashboardSheet.removeChart(chart));
    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    if (dashboardSheet && helperSheet) {
        updatePlatformDistributionChart(dashboardSheet, helperSheet);
        updateApplicationsOverTimeChart(dashboardSheet, helperSheet);
        updateApplicationFunnelChart(dashboardSheet, helperSheet);
    }
}


function updatePlatformDistributionChart(dashboardSheet, helperSheet) {
    const dataRange = helperSheet.getRange("A2:B");
    if (helperSheet.getRange("A2").isBlank()) return;

    const chart = dashboardSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(dataRange)
        .setPosition(13, 2, 0, 0) // Anchor at B13
        .setOption('title', "Platform Distribution")
        .setOption('pieHole', 0.4)
        .setOption('width', 480)
        .setOption('height', 300)
        .setOption('is3D', true)
        .setOption('legend', { position: 'right' })
        .setOption('colors', BRAND_COLORS_CHART_ARRAY())
        .build();
    dashboardSheet.insertChart(chart);
}

function updateApplicationsOverTimeChart(dashboardSheet, helperSheet) {
    const dataRange = helperSheet.getRange("D2:E");
    if (helperSheet.getRange("D2").isBlank()) return;

    const chart = dashboardSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(dataRange)
        .setPosition(13, 8, 0, 0) // Anchor at H13
        .setOption('title', "Applications Over Time (Weekly)")
        .setOption('width', 480)
        .setOption('height', 300)
        .setOption('hAxis', { title: 'Week Starting', format: 'M/d' })
        .setOption('vAxis', { title: 'Applications', viewWindow: { min: 0 } })
        .setOption('legend', { position: 'none' })
        .setOption('colors', [BRAND_COLORS.LAPIS_LAZULI])
        .setOption('curveType', 'function')
        .build();
    dashboardSheet.insertChart(chart);
}

function updateApplicationFunnelChart(dashboardSheet, helperSheet) {
    const dataRange = helperSheet.getRange("G2:H");
    if (helperSheet.getRange("G2").isBlank()) return;

    const chart = dashboardSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(dataRange)
        .setPosition(30, 2, 0, 0) // Anchor at B30
        .setOption('title', "Application Funnel (Peak Stages)")
        .setOption('width', 480)
        .setOption('height', 300)
        .setOption('hAxis', { slantedText: true, slantedTextAngle: 30 })
        .setOption('vAxis', { title: 'Applications', viewWindow: { min: 0 } })
        .setOption('legend', { position: 'none' })
        .setOption('colors', [BRAND_COLORS.CAROLINA_BLUE])
        .build();
    dashboardSheet.insertChart(chart);
}
