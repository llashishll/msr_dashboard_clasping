/**
 * @OnlyCurrentDoc
 * Reads normalized event data from "Dashboard" sheet, filters by month,
 * processes Sunday/Wednesday into pivoted tables (adjusting date headers +1 day),
 * processes Special Satsangs into a simple list, calculates averages,
 * applies custom sorting, serves a web app, and allows exporting filtered
 * display data to Excel. Defaults to current/latest month.
 */

// --- Configuration ---
const SHEET_NAME = "Dashboard";
const HEADER_ROWS = 1;
const SCRIPT_TIMEZONE = "GMT+5:30"; // <<< ENSURE THIS IS CORRECT (e.g., "Asia/Kolkata")

// Define column indices
const COL_DAY = 1;
const COL_DATE = 2;
const COL_CENTRE = 3;
const COL_SANGAT = 4;
const COL_MODE = 5;
const COL_NAME = 6;
const COL_ARRIVAL = 7;
const COL_SATSANG = 8;
const COL_SAINT = 9;
const COL_BOOK = 10;
const COL_START = 11;
const COL_END = 12;

// Column names for event details (used for keys and headers in PIVOTED tables)
const eventDetailColumns = [
  { key: "totalSangat", name: "Total Sangat", index: COL_SANGAT - 1 },
  { key: "mode", name: "Mode of Satsang", index: COL_MODE - 1 },
  { key: "name", name: "Name", index: COL_NAME - 1 },
  { key: "arrival", name: "Arrival Time", index: COL_ARRIVAL - 1 },
  { key: "satsang", name: "Satsang", index: COL_SATSANG - 1 },
  { key: "saint", name: "Name of Saint", index: COL_SAINT - 1 },
  { key: "book", name: "Name of Book", index: COL_BOOK - 1 },
  { key: "start", name: "Start Time", index: COL_START - 1 },
  { key: "end", name: "End Time", index: COL_END - 1 },
];

// Headers/Keys for the SPECIAL SATSANG LIST TABLE (Order matters)
const specialEventHeaders = [
  { key: "date", name: "Date", index: COL_DATE - 1 }, // From Display Value
  { key: "day", name: "Day", index: COL_DAY - 1 }, // From Display Value
  { key: "centre", name: "Centre", index: COL_CENTRE - 1 }, // From Display Value
  // Include details from eventDetailColumns
  ...eventDetailColumns, // Spread operator to include all detail columns here
];

// Headers for Excel export (using original column names from sheet)
const excelHeaders = [
  "Day",
  "Date",
  "Centre",
  "Total Sangat",
  "Mode of Satsang",
  "Name",
  "Arrival Time",
  "Satsang",
  "Name of Saint",
  "Name of Book",
  "Start Time",
  "End Time",
];

// Custom Centre Ordering and Bolding
const customCentreOrder = [
  "ALWAR",
  "ALWAR-2",
  "BEHROR",
  "BHIWADI",
  "BHOJRAJKA",
  "CHIKANI",
  "DESULA",
  "FATEHPUR",
  "GOBINDGARH",
  "GUNTASHAPUR",
  "HALDEENA",
  "HAZIPUR",
  "JHALATALA",
  "KARANA",
  "KARNIKOT",
  "KASBA DEHRA",
  "KAYSA",
  "KHAIRTHAL",
  "KISHANGARH BAS",
  "LAXMANGARH",
  "MUBARIKPUR",
  "MUNDAWAR",
  "PAHADI",
  "PARVENI",
  "PAWTI",
  "PEPEAL KHERA",
  "PRATAPGARH",
  "RAJGARH",
  "RAMGARH",
  "RATA KHURD",
  "SHAHBAD",
  "SHAJHANPUR",
  "TAPUKARA",
  "TATARPUR",
];
const boldCentres = ["ALWAR", "ALWAR-2", "RAMGARH", "KHAIRTHAL", "CHIKANI"];

// --- End Configuration ---

/**
 * Main function to serve the HTML page for the web app.
 */
function doGet(e) {
  // Determine Sunday sort preference from URL parameter, default to 'alphabetical'
  const sundaySortPref =
    e && e.parameter && e.parameter.sundaySort
      ? e.parameter.sundaySort
      : "alphabetical";
  // Set script timezone early
  try {
    Session.setScriptTimeZone(SCRIPT_TIMEZONE);
  } catch (tzErr) {
    Logger.log("Could not set script timezone: " + tzErr);
  }
  const scriptTimeZone = Session.getScriptTimeZone();
  Logger.log("Script timezone set to: " + scriptTimeZone);

  const requestedMonth =
    e && e.parameter && e.parameter.month ? e.parameter.month : null;
  const template = HtmlService.createTemplateFromFile("Index");

  // --- Get Processed Data ---
  const processedDataResult = getDataForDashboard(
    requestedMonth,
    sundaySortPref
  );

  // --- Assign data to template ---
  template.sundayData = processedDataResult.sunday;
  template.wednesdayData = processedDataResult.wednesday;
  template.specialData = processedDataResult.special; // Array of objects or empty array
  template.error = processedDataResult.error;
  template.availableMonths = processedDataResult.availableMonths;
  template.selectedMonth = processedDataResult.selectedMonth;
  template.sundaySortPreference = sundaySortPref; // Pass current sort preference to the template

  // --- Assign Headers/Keys needed for rendering ---
  // For Pivoted Tables (Sun/Wed)
  template.eventDetailHeaders = eventDetailColumns.map((col) => col.name);
  template.eventDetailKeys = eventDetailColumns.map((col) => col.key);
  // For Special Satsang List Table
  template.specialEventTableHeaders = specialEventHeaders.map((h) => h.name); // Uses global const
  template.specialEventTableKeys = specialEventHeaders.map((h) => h.key); // Uses global const

  // --- Evaluate ---
  return template
    .evaluate()
    .setTitle("Satsang Dashboard by Day")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper function to parse a date value from the sheet's *underlying value*.
 * Returns date in 'YYYY-MM-DD' format *respecting SCRIPT_TIMEZONE*, or null if invalid.
 */
function parseDateValue(dateValue, rowIndex) {
  let parsedDate = null;
  const origin = `(Row ${rowIndex + 1}, Val: ${dateValue})`;

  // Try Date Object
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
    parsedDate = dateValue;
  }
  // Try Serial Number
  else if (typeof dateValue === "number" && dateValue > 0) {
    try {
      const baseDate = new Date(1899, 11, 30);
      const dateMillis = baseDate.getTime() + dateValue * 24 * 60 * 60 * 1000;
      const potentialDate = new Date(dateMillis);
      const offset = potentialDate.getTimezoneOffset() * 60 * 1000;
      const adjustedDate = new Date(dateMillis - offset);
      if (!isNaN(adjustedDate.getTime())) {
        parsedDate = adjustedDate;
      } else {
        Logger.log(`Failed serial date ${dateValue} ${origin}`);
      }
    } catch (dateErr) {
      Logger.log(`Error serial date ${dateValue} ${origin}: ${dateErr}`);
    }
  }
  // Try String Parsing
  else if (typeof dateValue === "string" && dateValue.trim() !== "") {
    const trimmedDate = dateValue.trim();
    try {
      // Try dd-MM-yyyy format
      let p = Utilities.parseDate(trimmedDate, SCRIPT_TIMEZONE, "dd-MM-yyyy");
      if (!isNaN(p.getTime())) {
        parsedDate = p;
      } else {
        // Try yyyy-MM-dd format
        try {
          p = Utilities.parseDate(trimmedDate, SCRIPT_TIMEZONE, "yyyy-MM-dd");
          if (!isNaN(p.getTime())) {
            parsedDate = p;
          } else {
            // Try dd/MM/yyyy format
            try {
              p = Utilities.parseDate(
                trimmedDate,
                SCRIPT_TIMEZONE,
                "dd/MM/yyyy"
              );
              if (!isNaN(p.getTime())) {
                parsedDate = p;
              } else {
                // Try MM/dd/yyyy format
                try {
                  p = Utilities.parseDate(
                    trimmedDate,
                    SCRIPT_TIMEZONE,
                    "MM/dd/yyyy"
                  );
                  if (!isNaN(p.getTime())) {
                    parsedDate = p;
                  } else {
                    // Try generic Date parsing
                    try {
                      const genericDate = new Date(trimmedDate);
                      if (!isNaN(genericDate.getTime())) {
                        parsedDate = genericDate;
                      } else {
                        Logger.log(
                          `Failed all string parse attempts ${origin}`
                        );
                      }
                    } catch (genericErr) {
                      Logger.log(
                        `Error generic date parsing ${trimmedDate} ${origin}: ${genericErr}`
                      );
                    }
                  }
                } catch (mmddErr) {
                  Logger.log(
                    `Error MM/dd/yyyy parsing ${trimmedDate} ${origin}: ${mmddErr}`
                  );
                }
              }
            } catch (ddmmErr) {
              Logger.log(
                `Error dd/MM/yyyy parsing ${trimmedDate} ${origin}: ${ddmmErr}`
              );
            }
          }
        } catch (yyyymmErr) {
          Logger.log(
            `Error yyyy-MM-dd parsing ${trimmedDate} ${origin}: ${yyyymmErr}`
          );
        }
      }
    } catch (parseErr) {
      Logger.log(`Error parsing string ${trimmedDate} ${origin}: ${parseErr}`);
    }
  }

  // Format valid date using SCRIPT_TIMEZONE
  if (parsedDate && !isNaN(parsedDate.getTime())) {
    try {
      return Utilities.formatDate(parsedDate, SCRIPT_TIMEZONE, "yyyy-MM-dd");
    } catch (formatErr) {
      Logger.log(
        `Error formatting ${parsedDate} in ${SCRIPT_TIMEZONE} ${origin}: ${formatErr}`
      );
      return null;
    }
  }

  return null;
}

/**
 * Fetches data, filters, categorizes. Processes Sun/Wed with pivot, Special as list.
 */
function getDataForDashboard(requestedMonth) {
  const scriptTimeZone = Session.getScriptTimeZone();
  let sundayRows = [];
  let wednesdayRows = [];
  let specialRows = []; // Store { valueRow, displayRow, ymd }
  let availableMonthsSet = new Set();
  let error = null;
  let actualSelectedMonth = null;
  let allValidDataItems = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${SHEET_NAME}' not found!`);
    const range = sheet.getDataRange();
    const rawDataValues = range.getValues();
    const rawDataDisplayValues = range.getDisplayValues();
    if (rawDataValues.length <= HEADER_ROWS) {
      return {
        sunday: null,
        wednesday: null,
        special: [],
        error: "No data rows found.",
        availableMonths: [],
        selectedMonth: null,
      };
    }

    // --- First Pass: Validate Dates & Collect Months ---
    Logger.log("Starting first pass...");
    for (let i = HEADER_ROWS; i < rawDataValues.length; i++) {
      const dateValue = rawDataValues[i][COL_DATE - 1];
      const parsedYMD = parseDateValue(dateValue, i); // Timezone-aware YYYY-MM-DD
      if (parsedYMD) {
        availableMonthsSet.add(parsedYMD.substring(0, 7));
        allValidDataItems.push({
          valueRow: rawDataValues[i],
          displayRow: rawDataDisplayValues[i],
          ymd: parsedYMD,
        });
      }
    }
    Logger.log(`Found ${availableMonthsSet.size} months with data.`);

    // --- Determine Actual Month to Display ---
    const sortedAvailableMonths = Array.from(availableMonthsSet)
      .sort()
      .reverse();
    const currentMonth = Utilities.formatDate(
      new Date(),
      scriptTimeZone,
      "yyyy-MM"
    );
    Logger.log(
      `Current month (${scriptTimeZone}): ${currentMonth}, Requested: ${requestedMonth}`
    );
    if (requestedMonth && availableMonthsSet.has(requestedMonth)) {
      actualSelectedMonth = requestedMonth;
    } else if (availableMonthsSet.has(currentMonth)) {
      actualSelectedMonth = currentMonth;
    } else if (sortedAvailableMonths.length > 0) {
      actualSelectedMonth = sortedAvailableMonths[0];
    } else {
      Logger.log("No valid data found for any month.");
      return {
        sunday: null,
        wednesday: null,
        special: [],
        error: "No data with valid dates found.",
        availableMonths: [],
        selectedMonth: null,
      };
    }
    Logger.log(`Selected month for display: ${actualSelectedMonth}`);

    // --- Second Pass: Filter and Categorize ---
    const filteredDataItems = allValidDataItems.filter((item) =>
      item.ymd.startsWith(actualSelectedMonth)
    );
    Logger.log(
      `Found ${filteredDataItems.length} rows for month ${actualSelectedMonth}. Categorizing...`
    );
    filteredDataItems.forEach((item) => {
      const dayValue = (item.displayRow[COL_DAY - 1] || "")
        .toString()
        .trim()
        .toLowerCase();
      const centreName = item.displayRow[COL_CENTRE - 1];
      if (!centreName || centreName.trim() === "") return;
      if (dayValue === "sunday" || dayValue === "रविवार") sundayRows.push(item);
      else if (dayValue === "wednesday" || dayValue === "बुधवार")
        wednesdayRows.push(item);
      else if (dayValue !== "") specialRows.push(item);
    });
    Logger.log(
      `Categorized - Sun: ${sundayRows.length}, Wed: ${wednesdayRows.length}, Special: ${specialRows.length}`
    );

    // --- Process Data ---
    const sundayResult = processAndSortData(
      sundayRows,
      actualSelectedMonth,
      true
    ); // Pivot Sun, isSundayData = true
    const wednesdayResult = processAndSortData(
      wednesdayRows,
      actualSelectedMonth,
      false // Pivot Wed, isSundayData = false
    );
    const specialResultArray = processSpecialSatsangList(specialRows); // Unpivot Special

    return {
      sunday: sundayResult,
      wednesday: wednesdayResult,
      special: specialResultArray, // Array
      error: null,
      availableMonths: sortedAvailableMonths.reverse(), // Ascending for dropdown
      selectedMonth: actualSelectedMonth,
    };
  } catch (err) {
    Logger.log("Error in getDataForDashboard: " + err + " Stack: " + err.stack);
    const sortedMonths = Array.from(availableMonthsSet).sort().reverse();
    return {
      sunday: null,
      wednesday: null,
      special: [],
      error: "Failed to load data: " + err.message,
      availableMonths: sortedMonths.reverse(),
      selectedMonth: actualSelectedMonth,
    };
  }
}

/**
 * Helper function to pivot Sun/Wed data, calculate averages, apply custom sorting,
 * and adjust date headers +1 day for display.
 * Uses ymd (in SCRIPT_TIMEZONE) for internal logic, display values for cells.
 */
function processAndSortData(dataItems, selectedMonth, isSundayData = false) {
  const scriptTimeZone = Session.getScriptTimeZone();
  
  // Early exit if dataItems is not an array or is empty
  if (!Array.isArray(dataItems) || dataItems.length === 0) {
    if (!isSundayData) {
      return { templateData: [], sortedDates: [] };
    }
    // If it's Sunday data, we'll proceed with empty dataItems
    dataItems = [];
  }

  let processedData = {};
  let uniqueDateInfo = {}; // Stores { "dd-MM-yyyy": "yyyy-MM-dd" }

  // Process data items if they exist
  dataItems.forEach((item, index) => {
    const displayRow = item.displayRow;
    const valueRow = item.valueRow;
    const ymd = item.ymd;

    const centreName = displayRow[COL_CENTRE - 1]
      ? displayRow[COL_CENTRE - 1].toString().trim()
      : `_UnknownCentre_${index}`;
    let datePivotKey = "";
    let dateSortKey = "";

    if (ymd) {
      try {
        dateSortKey = ymd;

        let originalDate;
        try {
          originalDate = Utilities.parseDate(ymd, scriptTimeZone, "yyyy-MM-dd");
          if (isNaN(originalDate.getTime()))
            throw new Error("parseDate failed");
        } catch (parseUtilErr) {
          Logger.log(
            `Util.parseDate fail ymd=${ymd}, fallback new Date: ${parseUtilErr}`
          );
          const parts = ymd.split("-");
          const year = parseInt(parts[0], 10);
          const month = parseInt(parts[1], 10) - 1;
          const day = parseInt(parts[2], 10);
          originalDate = new Date(year, month, day);
          if (isNaN(originalDate.getTime())) throw new Error("Fallback fail");
        }

        // Removed +1 day adjustment here
        datePivotKey = Utilities.formatDate(
          originalDate,
          scriptTimeZone,
          "dd-MM-yyyy"
        );

        if (!uniqueDateInfo[datePivotKey]) {
          uniqueDateInfo[datePivotKey] = dateSortKey;
        }
      } catch (err) {
        Logger.log(
          `Row ${index}: Err adjust date ymd=${ymd}: ${err}. Using original.`
        );
        try {
          const parts = ymd.split("-");
          datePivotKey = `${parts[2]}-${parts[1]}-${parts[0]}`;
          if (!uniqueDateInfo[datePivotKey]) {
            uniqueDateInfo[datePivotKey] = ymd;
          }
        } catch (fallbackErr) {
          Logger.log(`Fail keys ymd=${ymd}`);
          return; // This return only exits the forEach callback
        }
      }
    } else {
      Logger.log(`Row ${index}: Missing YMD.`);
      return; // This return only exits the forEach callback
    }

    if (!processedData[centreName]) {
      processedData[centreName] = {
        dateData: {},
        totalSangatSum: 0,
        sangatCount: 0,
      };
    }

    const eventDetails = {};
    eventDetailColumns.forEach((col) => {
      let cellDisplayValue = displayRow[col.index];
      eventDetails[col.key] =
        cellDisplayValue === null || cellDisplayValue === undefined
          ? ""
          : cellDisplayValue;
    });

    if (!processedData[centreName].dateData[datePivotKey]) {
      processedData[centreName].dateData[datePivotKey] = eventDetails;

      const sangatValue = valueRow[COL_SANGAT - 1];
      let numericSangat = NaN;
      if (typeof sangatValue === "number") {
        numericSangat = sangatValue;
      } else if (typeof sangatValue === "string") {
        numericSangat = parseFloat(sangatValue.replace(/,/g, ""));
      }

      if (!isNaN(numericSangat)) {
        processedData[centreName].totalSangatSum += numericSangat;
        processedData[centreName].sangatCount++;
      }
    }
  });

  const sortedDateKeys = Object.keys(uniqueDateInfo).sort((a, b) => {
    const sortKeyA = uniqueDateInfo[a];
    const sortKeyB = uniqueDateInfo[b];
    return sortKeyA.localeCompare(sortKeyB);
  });

  Logger.log(
    `Sorted (original) date pivot keys ('dd-MM-yyyy') for month ${selectedMonth} (${
      isSundayData ? "Sunday" : "Other"
    }): ${sortedDateKeys.join(", ")}`
  );

  const finalTemplateData = [];
  const centresAlreadyAddedToFinal = new Set();

  // If it's Sunday data, iterate through customCentreOrder first to ensure all are present
  if (isSundayData) {
    customCentreOrder.forEach((centreName) => {
      const centreDataFromProcessed = processedData[centreName];
      let isMissingData = false;

      // Check for missing data only if there are actual Sundays in the month's data
      if (sortedDateKeys.length > 0) {
        if (!centreDataFromProcessed) {
          // Centre has no Sunday data entries at all this month
          isMissingData = true;
        } else {
          // Check if this centre is missing data for any specific Sunday date
          for (const dateKey of sortedDateKeys) {
            if (!centreDataFromProcessed.dateData[dateKey]) {
              isMissingData = true; // Missing data for at least one Sunday
              break;
            }
          }
        }
      }
      // If sortedDateKeys.length is 0 (no Sundays in the data for this month),
      // then isMissingData remains false, as it's not "missing" data for a non-existent Sunday.

      const averageSangat =
        centreDataFromProcessed && centreDataFromProcessed.sangatCount > 0
          ? (
              centreDataFromProcessed.totalSangatSum /
              centreDataFromProcessed.sangatCount
            ).toFixed(1)
          : "N/A";

      finalTemplateData.push({
        centre: centreName,
        isBold: boldCentres.includes(centreName),
        dateData: centreDataFromProcessed
          ? centreDataFromProcessed.dateData
          : {},
        averageSangat: averageSangat,
        isMissingData: isMissingData, // This flag will be used for highlighting
      });
      centresAlreadyAddedToFinal.add(centreName);
    });
  }

  // Add remaining centres from processedData (those not in customCentreOrder, or if not SundayData)
  // These are centres that had data but weren't in the custom order list (if isSundayData is false, all centres fall here)
  Object.keys(processedData)
    .sort((a, b) => a.localeCompare(b)) // Sort for consistent order
    .forEach((centreName) => {
      if (!centresAlreadyAddedToFinal.has(centreName)) {
        // Avoid duplicating if already added via customOrder
        const centreData = processedData[centreName];
        const averageSangat =
          centreData.sangatCount > 0
            ? (centreData.totalSangatSum / centreData.sangatCount).toFixed(1)
            : "N/A";
        finalTemplateData.push({
          centre: centreName,
          isBold: boldCentres.includes(centreName),
          dateData: centreData.dateData,
          averageSangat: averageSangat,
          isMissingData: false, // Default to false for centres not in custom list or for non-Sunday data
        });
      }
    });

  return {
    templateData: finalTemplateData,
    sortedDates: sortedDateKeys,
  };
}

/**
 * Processes Special Satsang rows into a simple sorted list of event objects.
 * @param {Array<Object>} dataItems Array of objects { valueRow, displayRow, ymd }.
 * @return {Array<Object>} Sorted array of event objects using display values.
 */
function processSpecialSatsangList(dataItems) {
  if (!dataItems || dataItems.length === 0) {
    return [];
  }
  Logger.log(
    `Processing ${dataItems.length} special satsang items into a list...`
  );
  const eventList = [];
  dataItems.forEach((item) => {
    const displayRow = item.displayRow;
    const ymd = item.ymd; // YYYY-MM-DD for sorting
    const eventObject = { ymdForSort: ymd }; // Add original date for sorting
    specialEventHeaders.forEach((headerInfo) => {
      // Use the defined headers config
      let cellDisplayValue = displayRow[headerInfo.index];
      eventObject[headerInfo.key] =
        cellDisplayValue === null || cellDisplayValue === undefined
          ? ""
          : cellDisplayValue;
    });
    eventList.push(eventObject);
  });
  // Sort the list by date (using ymdForSort)
  eventList.sort((a, b) => {
    if (!a.ymdForSort || !b.ymdForSort) return 0;
    return a.ymdForSort.localeCompare(b.ymdForSort); // Sort by YYYY-MM-DD string
  });
  Logger.log(
    `Finished processing special list. ${eventList.length} events found.`
  );
  return eventList;
}

/**
 * Fetches data filtered by month, creates Excel matching the dashboard web app format.
 */
function exportFilteredDataToExcel() {
  const scriptTimeZone = Session.getScriptTimeZone();
  const now = new Date();
  const selectedMonth = Utilities.formatDate(now, scriptTimeZone, "yyyy-MM"); // auto-detect current month
  Logger.log("Auto-detected month for export: " + selectedMonth);

  let spreadsheet = null;

  try {
    // Get data using the same process as the web app
    let sundayRows = [];
    let wednesdayRows = [];
    let specialRows = [];
    let allValidDataItems = [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(SHEET_NAME);
    if (!sourceSheet) throw new Error(`Sheet '${SHEET_NAME}' not found!`);

    const range = sourceSheet.getDataRange();
    const rawDataValues = range.getValues();
    const rawDataDisplayValues = range.getDisplayValues();

    if (rawDataValues.length <= HEADER_ROWS) {
      throw new Error("No data found.");
    }

    // First pass: Validate dates & collect
    for (let i = HEADER_ROWS; i < rawDataValues.length; i++) {
      const dateValue = rawDataValues[i][COL_DATE - 1];
      const parsedYMD = parseDateValue(dateValue, i); // Timezone-aware YYYY-MM-DD
      if (parsedYMD) {
        allValidDataItems.push({
          valueRow: rawDataValues[i],
          displayRow: rawDataDisplayValues[i],
          ymd: parsedYMD,
        });
      }
    }

    // Filter and categorize data for the selected month
    const filteredDataItems = allValidDataItems.filter((item) =>
      item.ymd.startsWith(selectedMonth)
    );
    Logger.log(
      `Export: Found ${filteredDataItems.length} valid rows for month ${selectedMonth}. Categorizing...`
    );

    if (filteredDataItems.length === 0) {
      throw new Error("No data found for current month: " + selectedMonth);
    }

    filteredDataItems.forEach((item) => {
      const dayValue = (item.displayRow[COL_DAY - 1] || "")
        .toString()
        .trim()
        .toLowerCase();
      const centreName = item.displayRow[COL_CENTRE - 1];
      if (!centreName || centreName.trim() === "") return;
      if (dayValue === "sunday" || dayValue === "रविवार") sundayRows.push(item);
      else if (dayValue === "wednesday" || dayValue === "बुधवार")
        wednesdayRows.push(item);
      else if (dayValue !== "") specialRows.push(item);
    });

    Logger.log(
      `Export: Categorized - Sun: ${sundayRows.length}, Wed: ${wednesdayRows.length}, Special: ${specialRows.length}`
    );

    // Process data using the same functions used by the web app
    const sundayResult = processAndSortData(sundayRows, selectedMonth, true); // true for Sunday data
    const wednesdayResult = processAndSortData(wednesdayRows, selectedMonth, false); // false for Wednesday data
    const specialResultArray = processSpecialSatsangList(specialRows);

    // Create Excel file
    const timestamp = Utilities.formatDate(
      new Date(),
      scriptTimeZone,
      "yyyyMMdd_HHmmss"
    );
    const newSheetName = `Satsang_Report_${selectedMonth}_${timestamp}`;
    spreadsheet = SpreadsheetApp.create(newSheetName);
    Logger.log("Created temp sheet: " + spreadsheet.getName());

    // Create pivoted Sunday sheet
    if (sundayResult.templateData.length > 0) {
      const sundaySheet = spreadsheet.insertSheet("Sunday");
      createPivotedExcelSheet(sundaySheet, sundayResult, "Sunday Satsang");
    }

    // Create pivoted Wednesday sheet
    if (wednesdayResult.templateData.length > 0) {
      const wednesdaySheet = spreadsheet.insertSheet("Wednesday");
      createPivotedExcelSheet(
        wednesdaySheet,
        wednesdayResult,
        "Wednesday Satsang"
      );
    }

    // Create Special Satsang sheet (list format)
    if (specialResultArray.length > 0) {
      const specialSheet = spreadsheet.insertSheet("Special Satsangs");
      createSpecialExcelSheet(specialSheet, specialResultArray);
    }

    // Delete default sheet if we created at least one data sheet
    if (
      sundayResult.templateData.length > 0 ||
      wednesdayResult.templateData.length > 0 ||
      specialResultArray.length > 0
    ) {
      const defaultSheet = spreadsheet.getSheetByName("Sheet1");
      if (defaultSheet) spreadsheet.deleteSheet(defaultSheet);
    } else {
      // If no data sheets were created, just use Sheet1 to hold a message
      const defaultSheet = spreadsheet.getSheetByName("Sheet1");
      if (defaultSheet) {
        defaultSheet
          .getRange("A1")
          .setValue("No data found for the selected month: " + selectedMonth);
        defaultSheet.autoResizeColumn(1);
      }
    }

    SpreadsheetApp.flush();

    // Export as Excel
    const spreadsheetId = spreadsheet.getId();
    const file = DriveApp.getFileById(spreadsheetId);

    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
      const options = {
        method: "GET",
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const blob = response.getBlob().setName(newSheetName + ".xlsx");
        const tempFolder = DriveApp.getRootFolder();
        const excelFile = tempFolder.createFile(blob);
        excelFile.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );
        const downloadUrl = `https://drive.google.com/uc?export=download&id=${excelFile.getId()}`;
        Logger.log(
          "Excel file created: " + excelFile.getName() + " URL: " + downloadUrl
        );

        PropertiesService.getScriptProperties().setProperty(
          "fileToDeleteId",
          spreadsheetId + "," + excelFile.getId()
        );

        let existingTrigger = false;
        ScriptApp.getProjectTriggers().forEach((trigger) => {
          if (trigger.getHandlerFunction() === "deleteTemporaryFile") {
            existingTrigger = true;
          }
        });

        if (!existingTrigger) {
          ScriptApp.newTrigger("deleteTemporaryFile")
            .timeBased()
            .after(10 * 60 * 1000)
            .create();
          Logger.log("Created cleanup trigger.");
        }

        return downloadUrl;
      } else {
        throw new Error(`Failed Excel export (HTTP ${responseCode}).`);
      }
    } catch (e) {
      Logger.log("Err conversion/sharing/cleanup: " + e);
      if (spreadsheet) {
        try {
          DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
        } catch (trashErr) {}
      }
      throw e;
    }
  } catch (err) {
    Logger.log("Err exportFilteredData: " + err + " Stack: " + err.stack);
    if (spreadsheet) {
      try {
        DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
      } catch (trashErr) {}
    }
    throw new Error("Export failed: " + err.message);
  }
}

/**
 * Helper function to create a pivoted Excel sheet (for Sunday and Wednesday)
 * @param {Sheet} sheet The sheet to populate
 * @param {Object} result The result object from processAndSortData
 * @param {String} title The title to display at the top of the sheet
 */
function createPivotedExcelSheet(sheet, result, title) {
  // Add title
  sheet.getRange("A1").setValue(title);
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");

  // Create header row with dates
  const datesHeaderRow = 3;
  sheet.getRange(datesHeaderRow, 1).setValue("Centre");
  sheet.getRange(datesHeaderRow, 2).setValue("Avg. Sangat");

  let columnIndex = 3;
  result.sortedDates.forEach((dateStr) => {
    sheet.getRange(datesHeaderRow, columnIndex).setValue(dateStr);
    columnIndex++;
  });

  // Apply header formatting
  const headerRange = sheet.getRange(datesHeaderRow, 1, 1, columnIndex - 1);
  headerRange.setFontWeight("bold").setBackground("#f3f3f3");

  // Add data rows
  let rowIndex = datesHeaderRow + 1;
  result.templateData.forEach((centreData) => {
    // Centre name and average
    sheet.getRange(rowIndex, 1).setValue(centreData.centre);
    if (centreData.isBold) {
      sheet.getRange(rowIndex, 1).setFontWeight("bold");
    }
    if (centreData.isMissingData) {
      sheet.getRange(rowIndex, 1).setFontColor("#FF0000"); // Red for missing data
    }
    sheet.getRange(rowIndex, 2).setValue(centreData.averageSangat);

    // Date data
    columnIndex = 3;
    result.sortedDates.forEach((dateStr) => {
      if (centreData.dateData[dateStr]) {
        const cellData = centreData.dateData[dateStr].totalSangat || "";
        sheet.getRange(rowIndex, columnIndex).setValue(cellData);
      }
      columnIndex++;
    });

    rowIndex++;
  });

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, columnIndex - 1);
}

/**
 * Helper function to create a special satsang Excel sheet in list format
 * @param {Sheet} sheet The sheet to populate
 * @param {Array} specialEvents Array of special event objects
 */
function createSpecialExcelSheet(sheet, specialEvents) {
  // Add title
  sheet.getRange("A1").setValue("Special Satsang Events");
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");

  // Create header row
  const headerRow = 3;
  let columnIndex = 1;

  specialEventHeaders.forEach((header) => {
    sheet.getRange(headerRow, columnIndex).setValue(header.name);
    columnIndex++;
  });

  // Apply header formatting
  const headerRange = sheet.getRange(
    headerRow,
    1,
    1,
    specialEventHeaders.length
  );
  headerRange.setFontWeight("bold").setBackground("#f3f3f3");

  // Add data rows
  let rowIndex = headerRow + 1;
  specialEvents.forEach((event) => {
    columnIndex = 1;
    specialEventHeaders.forEach((header) => {
      sheet.getRange(rowIndex, columnIndex).setValue(event[header.key] || "");
      columnIndex++;
    });
    rowIndex++;
  });

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, specialEventHeaders.length);
}

/**
 * Function to delete temporary files created during Excel export
 */
function deleteTemporaryFile() {
  try {
    const prop =
      PropertiesService.getScriptProperties().getProperty("fileToDeleteId");
    if (!prop) return;

    const fileIds = prop.split(",");
    fileIds.forEach((fileId) => {
      try {
        const file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
        Logger.log("Deleted temporary file: " + fileId);
      } catch (e) {
        Logger.log("Could not delete file " + fileId + ": " + e);
      }
    });

    PropertiesService.getScriptProperties().deleteProperty("fileToDeleteId");
  } catch (e) {
    Logger.log("Error in cleanup: " + e);
  }
}

/**
 * Function to get the filtered data for client-side use
 */
function getFilteredDataForExport(requestedMonth) {
  try {
    const result = getDataForDashboard(requestedMonth);
    return result;
  } catch (e) {
    Logger.log("Error getting filtered data: " + e);
    return { error: "Failed to get filtered data: " + e.message };
  }
}

/**
 * Exports filtered data to Excel for the specified month
 * This function is called by the regular Export button
 */
function exportFilteredDataToExcel(requestedMonth) {
  const scriptTimeZone = Session.getScriptTimeZone();
  const selectedMonth =
    requestedMonth ||
    Utilities.formatDate(new Date(), scriptTimeZone, "yyyy-MM");
  Logger.log("Exporting filtered data for month: " + selectedMonth);

  let spreadsheet = null;

  try {
    // Get data using the same process as the web app
    const processedDataResult = getDataForDashboard(selectedMonth);

    if (processedDataResult.error) {
      throw new Error(processedDataResult.error);
    }

    const sundayResult = processedDataResult.sunday;
    const wednesdayResult = processedDataResult.wednesday;

    // Create Excel file
    const timestamp = Utilities.formatDate(
      new Date(),
      scriptTimeZone,
      "yyyyMMdd_HHmmss"
    );
    const newSheetName = `Satsang_Report_${selectedMonth}_${timestamp}`;
    spreadsheet = SpreadsheetApp.create(newSheetName);
    Logger.log("Created temp sheet: " + spreadsheet.getName());

    // Create pivoted Sunday sheet
    if (
      sundayResult &&
      sundayResult.templateData &&
      sundayResult.templateData.length > 0
    ) {
      const sundaySheet = spreadsheet.insertSheet("Sunday");
      createPivotedExcelSheet(sundaySheet, sundayResult, "Sunday Satsang");
    }

    // Create pivoted Wednesday sheet
    if (
      wednesdayResult &&
      wednesdayResult.templateData &&
      wednesdayResult.templateData.length > 0
    ) {
      const wednesdaySheet = spreadsheet.insertSheet("Wednesday");
      createPivotedExcelSheet(
        wednesdaySheet,
        wednesdayResult,
        "Wednesday Satsang"
      );
    }

    // Delete default sheet if we created at least one data sheet
    if (
      (sundayResult &&
        sundayResult.templateData &&
        sundayResult.templateData.length > 0) ||
      (wednesdayResult &&
        wednesdayResult.templateData &&
        wednesdayResult.templateData.length > 0)
    ) {
      const defaultSheet = spreadsheet.getSheetByName("Sheet1");
      if (defaultSheet) spreadsheet.deleteSheet(defaultSheet);
    } else {
      // If no data sheets were created, just use Sheet1 to hold a message
      const defaultSheet = spreadsheet.getSheetByName("Sheet1");
      if (defaultSheet) {
        defaultSheet
          .getRange("A1")
          .setValue("No data found for the selected month: " + selectedMonth);
        defaultSheet.autoResizeColumn(1);
      }
    }

    SpreadsheetApp.flush();

    // Export as Excel
    const spreadsheetId = spreadsheet.getId();
    const file = DriveApp.getFileById(spreadsheetId);

    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
      const options = {
        method: "GET",
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const blob = response.getBlob().setName(newSheetName + ".xlsx");
        const tempFolder = DriveApp.getRootFolder();
        const excelFile = tempFolder.createFile(blob);
        excelFile.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );
        const downloadUrl = `https://drive.google.com/uc?export=download&id=${excelFile.getId()}`;
        Logger.log(
          "Excel file created: " + excelFile.getName() + " URL: " + downloadUrl
        );

        // Add to cleanup list
        const prop =
          PropertiesService.getScriptProperties().getProperty(
            "fileToDeleteId"
          ) || "";
        const newProp = prop
          ? prop + "," + spreadsheetId + "," + excelFile.getId()
          : spreadsheetId + "," + excelFile.getId();
        PropertiesService.getScriptProperties().setProperty(
          "fileToDeleteId",
          newProp
        );

        let existingTrigger = false;
        ScriptApp.getProjectTriggers().forEach((trigger) => {
          if (trigger.getHandlerFunction() === "deleteTemporaryFile") {
            existingTrigger = true;
          }
        });

        if (!existingTrigger) {
          ScriptApp.newTrigger("deleteTemporaryFile")
            .timeBased()
            .after(10 * 60 * 1000)
            .create();
          Logger.log("Created cleanup trigger.");
        }

        return downloadUrl;
      } else {
        throw new Error(`Failed Excel export (HTTP ${responseCode}).`);
      }
    } catch (e) {
      Logger.log("Err conversion/sharing/cleanup: " + e);
      if (spreadsheet) {
        try {
          DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
        } catch (trashErr) {}
      }
      throw e;
    }
  } catch (err) {
    Logger.log("Err exportFilteredData: " + err + " Stack: " + err.stack);
    if (spreadsheet) {
      try {
        DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
      } catch (trashErr) {}
    }
    throw new Error("Export failed: " + err.message);
  }
}

/**
 * Include HTML content from the specified file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Returns the available months for the UI dropdown
 */
function getAvailableMonths() {
  try {
    const allData = getDataForDashboard(null);
    return {
      months: allData.availableMonths,
      selectedMonth: allData.selectedMonth,
    };
  } catch (e) {
    Logger.log("Error getting months: " + e);
    return {
      months: [],
      selectedMonth: null,
      error: "Failed to get months: " + e.message,
    };
  }
}

/**
 * Function to handle month change from the client-side
 */
function changeMonth(newMonth) {
  try {
    return getDataForDashboard(newMonth);
  } catch (e) {
    Logger.log("Error changing month: " + e);
    return { error: "Failed to change month: " + e.message };
  }
}

/**
 * Exports complete data with all columns in the same format as the web app
 */
function findMissingDateEntries(selectedMonth) {
  const scriptTimeZone = Session.getScriptTimeZone();
  const currentMonth =
    selectedMonth ||
    Utilities.formatDate(new Date(), scriptTimeZone, "yyyy-MM");
  const processedDataResult = getDataForDashboard(currentMonth);

  if (processedDataResult.error) {
    throw new Error(processedDataResult.error);
  }

  const results = [];

  // Check Sundays
  const sundayResult = processedDataResult.sunday;
  if (
    sundayResult &&
    sundayResult.templateData &&
    sundayResult.templateData.length > 0
  ) {
    const sundayDates = sundayResult.sortedDates || [];
    sundayResult.templateData.forEach((center) => {
      const centerName = center.centre;
      const missingDates = sundayDates.filter((date) => !center.dateData[date]);
      if (missingDates.length > 0) {
        results.push(`${centerName} (Sunday): ${missingDates.join(", ")}`);
      }
    });
  }

  // Check Wednesdays
  const wednesdayResult = processedDataResult.wednesday;
  const staticWednesdayCenters = [
    "ALWAR", "ALWAR-2", "BEHROR", "BHIWADI", "CHIKANI", "FATEHPUR", "GOBINDGARH",
    "HALDEENA", "HAZIPUR", "JHALATALA", "KARANA", "KARNIKOT", "KHAIRTHAL",
    "KISHANGARH BAS", "LAXMANGARH", "MUBARIKPUR", "PAWTI", "PEPEAL KHERA",
    "RAJGARH", "RAMGARH", "RATA KHURD", "SHAJHANPUR"
  ];
  if (
    wednesdayResult &&
    wednesdayResult.sortedDates &&
    wednesdayResult.sortedDates.length > 0
  ) {
    const wednesdayDates = wednesdayResult.sortedDates || [];
    staticWednesdayCenters.forEach(centerName => {
      // Find the center in the data
      const center = (wednesdayResult.templateData || []).find(row => row.centre === centerName);
      let missingDates = [];
      if (!center) {
        // If the center has no data at all, all dates are missing
        missingDates = wednesdayDates.slice();
      } else {
        // Otherwise, check which dates are missing
        missingDates = wednesdayDates.filter(date => !center.dateData[date]);
      }
      if (missingDates.length > 0) {
        results.push(`${centerName} (Wednesday): ${missingDates.join(", ")}`);
      }
    });
  }

  return results;
}

function exportCompleteDataToExcel(requestedMonth) {
  const scriptTimeZone = Session.getScriptTimeZone();
  const selectedMonth =
    requestedMonth ||
    Utilities.formatDate(new Date(), scriptTimeZone, "yyyy-MM");
  Logger.log("Exporting complete data for month: " + selectedMonth);

  let spreadsheet = null;

  try {
    // Get data using the same process as the web app
    const processedDataResult = getDataForDashboard(selectedMonth);

    if (processedDataResult.error) {
      throw new Error(processedDataResult.error);
    }

    const sundayResult = processedDataResult.sunday;
    const wednesdayResult = processedDataResult.wednesday;
    const specialResultArray = processedDataResult.special;

    // Create Excel file
    const timestamp = Utilities.formatDate(
      new Date(),
      scriptTimeZone,
      "yyyyMMdd_HHmmss"
    );
    const newSheetName = `Complete_Satsang_Report_${selectedMonth}_${timestamp}`;
    spreadsheet = SpreadsheetApp.create(newSheetName);
    Logger.log("Created temp sheet: " + spreadsheet.getName());

    // Create pivoted Sunday sheet
    if (
      sundayResult &&
      sundayResult.templateData &&
      sundayResult.templateData.length > 0
    ) {
      const sundaySheet = spreadsheet.insertSheet("Sunday");
      createCompletePivotedExcelSheet(
        sundaySheet,
        sundayResult,
        "Sunday Satsang"
      );
    }

    // Create pivoted Wednesday sheet
    if (
      wednesdayResult &&
      wednesdayResult.templateData &&
      wednesdayResult.templateData.length > 0
    ) {
      const wednesdaySheet = spreadsheet.insertSheet("Wednesday");
      createCompletePivotedExcelSheet(
        wednesdaySheet,
        wednesdayResult,
        "Wednesday Satsang"
      );
    }

    // Create Special Satsang sheet (list format)
    if (specialResultArray && specialResultArray.length > 0) {
      const specialSheet = spreadsheet.insertSheet("Special Satsangs");
      createCompleteSpecialExcelSheet(specialSheet, specialResultArray);
    }

    // Delete default sheet if we created at least one data sheet
    if (
      (sundayResult &&
        sundayResult.templateData &&
        sundayResult.templateData.length > 0) ||
      (wednesdayResult &&
        wednesdayResult.templateData &&
        wednesdayResult.templateData.length > 0) ||
      (specialResultArray && specialResultArray.length > 0)
    ) {
      const defaultSheet = spreadsheet.getSheetByName("Sheet1");
      if (defaultSheet) spreadsheet.deleteSheet(defaultSheet);
    } else {
      // If no data sheets were created, just use Sheet1 to hold a message
      const defaultSheet = spreadsheet.getSheetByName("Sheet1");
      if (defaultSheet) {
        defaultSheet
          .getRange("A1")
          .setValue("No data found for the selected month: " + selectedMonth);
        defaultSheet.autoResizeColumn(1);
      }
    }

    SpreadsheetApp.flush();

    // Export as Excel
    const spreadsheetId = spreadsheet.getId();
    const file = DriveApp.getFileById(spreadsheetId);

    try {
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
      const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx`;
      const options = {
        method: "GET",
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true,
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode === 200) {
        const blob = response.getBlob().setName(newSheetName + ".xlsx");
        const tempFolder = DriveApp.getRootFolder();
        const excelFile = tempFolder.createFile(blob);
        excelFile.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );
        const downloadUrl = `https://drive.google.com/uc?export=download&id=${excelFile.getId()}`;
        Logger.log(
          "Complete Excel file created: " +
            excelFile.getName() +
            " URL: " +
            downloadUrl
        );

        // Add to cleanup list
        const prop =
          PropertiesService.getScriptProperties().getProperty(
            "fileToDeleteId"
          ) || "";
        const newProp = prop
          ? prop + "," + spreadsheetId + "," + excelFile.getId()
          : spreadsheetId + "," + excelFile.getId();
        PropertiesService.getScriptProperties().setProperty(
          "fileToDeleteId",
          newProp
        );

        let existingTrigger = false;
        ScriptApp.getProjectTriggers().forEach((trigger) => {
          if (trigger.getHandlerFunction() === "deleteTemporaryFile") {
            existingTrigger = true;
          }
        });

        if (!existingTrigger) {
          ScriptApp.newTrigger("deleteTemporaryFile")
            .timeBased()
            .after(10 * 60 * 1000)
            .create();
          Logger.log("Created cleanup trigger.");
        }

        return downloadUrl;
      } else {
        throw new Error(`Failed Excel export (HTTP ${responseCode}).`);
      }
    } catch (e) {
      Logger.log("Err conversion/sharing/cleanup: " + e);
      if (spreadsheet) {
        try {
          DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
        } catch (trashErr) {}
      }
      throw e;
    }
  } catch (err) {
    Logger.log("Err exportCompleteData: " + err + " Stack: " + err.stack);
    if (spreadsheet) {
      try {
        DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
      } catch (trashErr) {}
    }
    throw new Error("Complete export failed: " + err.message);
  }
}

/**
 * Helper function to create a complete pivoted Excel sheet with all columns
 */
function createCompletePivotedExcelSheet(sheet, result, title) {
  // Add title
  sheet.getRange("A1").setValue(title);
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");

  // Create header row with dates
  const datesHeaderRow = 3;
  sheet.getRange(datesHeaderRow, 1).setValue("Centre");
  sheet.getRange(datesHeaderRow, 2).setValue("Avg. Sangat");

  let columnIndex = 3;
  result.sortedDates.forEach((dateStr) => {
    // For each date, we'll create multiple columns (one for each event detail)
    const dateColStart = columnIndex;
    sheet
      .getRange(datesHeaderRow, columnIndex, 1, eventDetailColumns.length)
      .merge();
    sheet.getRange(datesHeaderRow, columnIndex).setValue(dateStr);
    sheet
      .getRange(datesHeaderRow, columnIndex)
      .setHorizontalAlignment("center");

    // Add sub-headers for each column
    const subHeaderRow = datesHeaderRow + 1;
    eventDetailColumns.forEach((col) => {
      sheet.getRange(subHeaderRow, columnIndex).setValue(col.name);
      columnIndex++;
    });
  });

  // Apply header formatting
  const headerRange = sheet.getRange(datesHeaderRow, 1, 1, columnIndex - 1);
  headerRange.setFontWeight("bold").setBackground("#f3f3f3");

  const subHeaderRange = sheet.getRange(
    datesHeaderRow + 1,
    3,
    1,
    columnIndex - 3
  );
  subHeaderRange.setFontWeight("bold").setBackground("#e6e6e6");

  // Add data rows
  let rowIndex = datesHeaderRow + 2; // Start after the main header and sub-header rows
  result.templateData.forEach((centreData) => {
    // Centre name and average
    sheet.getRange(rowIndex, 1).setValue(centreData.centre);
    if (centreData.isBold) {
      sheet.getRange(rowIndex, 1).setFontWeight("bold");
    }
    if (centreData.isMissingData) {
      sheet.getRange(rowIndex, 1).setFontColor("#FF0000"); // Red for missing data
    }
    sheet.getRange(rowIndex, 2).setValue(centreData.averageSangat);

    // Date data with all columns
    columnIndex = 3;
    result.sortedDates.forEach((dateStr) => {
      if (centreData.dateData[dateStr]) {
        const eventData = centreData.dateData[dateStr];
        eventDetailColumns.forEach((col) => {
          const cellData = eventData[col.key] || "";
          sheet.getRange(rowIndex, columnIndex).setValue(cellData);
          columnIndex++;
        });
      } else {
        // Skip empty dates
        columnIndex += eventDetailColumns.length;
      }
    });

    rowIndex++;
  });

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, columnIndex - 1);

  // Freeze the header rows and first two columns
  sheet.setFrozenRows(datesHeaderRow + 1);
  sheet.setFrozenColumns(2);
}

/**
 * Helper function to create a complete special satsang Excel sheet with all columns
 */
function createCompleteSpecialExcelSheet(sheet, specialEvents) {
  // Add title
  sheet.getRange("A1").setValue("Special Satsang Events");
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");

  // Create header row
  const headerRow = 3;
  let columnIndex = 1;

  specialEventHeaders.forEach((header) => {
    sheet.getRange(headerRow, columnIndex).setValue(header.name);
    columnIndex++;
  });

  // Apply header formatting
  const headerRange = sheet.getRange(
    headerRow,
    1,
    1,
    specialEventHeaders.length
  );
  headerRange.setFontWeight("bold").setBackground("#f3f3f3");

  // Add data rows
  let rowIndex = headerRow + 1;
  specialEvents.forEach((event) => {
    columnIndex = 1;
    specialEventHeaders.forEach((header) => {
      sheet.getRange(rowIndex, columnIndex).setValue(event[header.key] || "");
      columnIndex++;
    });
    rowIndex++;
  });

  // Auto-resize columns for better readability
  sheet.autoResizeColumns(1, specialEventHeaders.length);

  // Freeze the header row and first three columns
  sheet.setFrozenRows(headerRow);
  sheet.setFrozenColumns(3);
}
