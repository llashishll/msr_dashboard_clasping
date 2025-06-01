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
 * Prioritizes 'yyyy-MM-dd', then attempts other common formats ('dd-MM-yyyy', 'dd/MM/yyyy', 'MM/dd/yyyy'),
 * then JS Date objects, then serial numbers. Logs warnings for non-standard formats.
 * Returns date in 'YYYY-MM-DD' format *respecting SCRIPT_TIMEZONE*, or null if invalid.
 */
function parseDateValue(inputValue, sheetRowIndex) {
  let parsedDateObj = null;
  const originalValueForLog = `(Row ${sheetRowIndex + 1}, Original Val: '${inputValue}')`;
  const scriptTimeZone = Session.getScriptTimeZone();

  const preferredFormat = "yyyy-MM-dd";
  const alternativeFormats = ["dd-MM-yyyy", "dd/MM/yyyy", "MM/dd/yyyy"];

  if (typeof inputValue === "string") {
    const trimmedDateStr = inputValue.trim();
    if (trimmedDateStr === "") {
      return null; // Empty string is not a valid date
    }

    // Attempt 1: Preferred format
    try {
      let d = Utilities.parseDate(trimmedDateStr, scriptTimeZone, preferredFormat);
      if (!isNaN(d.getTime())) {
        parsedDateObj = d;
      }
    } catch (e) { /* Ignore parse error, try next format */ }

    // Attempt 2: Alternative string formats
    if (!parsedDateObj) {
      for (const format of alternativeFormats) {
        try {
          let d = Utilities.parseDate(trimmedDateStr, scriptTimeZone, format);
          if (!isNaN(d.getTime())) {
            parsedDateObj = d;
            Logger.log(`parseDateValue: Parsed date using non-standard format '${format}'. ${originalValueForLog}. Recommend changing to '${preferredFormat}' in sheet.`);
            break;
          }
        } catch (e) { /* Ignore parse error, try next format */ }
      }
    }
  }

  // Attempt 3: JavaScript Date object
  if (!parsedDateObj && inputValue instanceof Date && !isNaN(inputValue.getTime())) {
    parsedDateObj = inputValue;
    Logger.log(`parseDateValue: Used JS Date object directly. ${originalValueForLog}. Recommend formatting as '${preferredFormat}' string in sheet.`);
  }

  // Attempt 4: Numeric serial date
  if (!parsedDateObj && typeof inputValue === 'number' && inputValue > 0) {
    // Basic check: very small numbers are unlikely to be valid serial dates for typical ranges.
    // Excel serial numbers for 2000-2050 are roughly 36526 - 54786.
    // Google Sheets might use a different base for serial numbers if not imported from Excel.
    // This check helps avoid misinterpreting regular numbers as dates.
    if (inputValue < 20000 && inputValue > 70000 && inputValue !== Math.floor(inputValue)) { // Heuristic: if it's not an integer or outside a broad excel range
         Logger.log(`parseDateValue: Numeric value ${inputValue} is unlikely a serial date. ${originalValueForLog}. Skipping serial date parsing.`);
    } else {
        try {
            // The common epoch for Excel serial dates is Dec 30, 1899 (for day 1 = Jan 1, 1900, due to Lotus 1-2-3 leap year bug)
            // or sometimes Jan 1, 1904 (Mac Excel default). Utilities.formatDate can often handle serial numbers directly
            // if the spreadsheet itself interprets them as dates, but here inputValue is a raw number.
            const baseDate = new Date(1899, 11, 30); // For Excel Windows epoch
            let dateMillis = baseDate.getTime() + (inputValue * 24 * 60 * 60 * 1000);

            // Check if this resulted in a date far in the past or future, indicating wrong epoch or not a date
            const tempDate = new Date(dateMillis);
            const year = tempDate.getFullYear();
            if (year < 1900 || year > 2100) { // Heuristic for typical date ranges
                 // Try adjusting for potential direct milliseconds if it's a huge number (less likely from sheets)
                if (inputValue > 1000000000000 && inputValue < 9999999999999) { // Plausible millisecond timestamp range
                    dateMillis = inputValue;
                     Logger.log(`parseDateValue: Interpreting numeric value as direct milliseconds. ${originalValueForLog}.`);
                } else {
                  throw new Error("Serial number resulted in unlikely year: " + year);
                }
            }

            const potentialDate = new Date(dateMillis);
            if (!isNaN(potentialDate.getTime())) {
                parsedDateObj = potentialDate;
                Logger.log(`parseDateValue: Parsed as a numeric serial date. ${originalValueForLog}. Recommend formatting as '${preferredFormat}' string in sheet.`);
            } else {
                Logger.log(`parseDateValue: Failed to parse serial number ${inputValue} into a valid date. ${originalValueForLog}`);
            }
        } catch (dateErr) {
            Logger.log(`parseDateValue: Error parsing numeric value ${inputValue} as serial date. ${originalValueForLog}: ${dateErr.message}`);
        }
    }
  }

  // Final logging for unparsed strings
  if (!parsedDateObj && typeof inputValue === 'string' && inputValue.trim() !== '') {
    Logger.log(`parseDateValue: Failed to parse date string '${inputValue.trim()}' after trying all formats. ${originalValueForLog}. Please use '${preferredFormat}' or one of the supported alternative formats.`);
  } else if (!parsedDateObj && inputValue && typeof inputValue !== 'string' && !(inputValue instanceof Date) && typeof inputValue !== 'number') {
    Logger.log(`parseDateValue: Unparseable date value of type ${typeof inputValue}. ${originalValueForLog}.`);
  }


  // Format valid date object to 'yyyy-MM-dd' string
  if (parsedDateObj && !isNaN(parsedDateObj.getTime())) {
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
    const wednesdayStaticCenters = getStaticWednesdayCentersFromSheet();

    const sundayResult = processAndSortData(
      sundayRows || [],
      actualSelectedMonth,
      true, // isSundayData = true
      customCentreOrder // allRelevantCenters for Sunday
    );
    const wednesdayResult = processAndSortData(
      wednesdayRows || [],
      actualSelectedMonth,
      false, // isSundayData = false
      wednesdayStaticCenters // allRelevantCenters for Wednesday
    );
    const specialResultArray = processSpecialSatsangList(specialRows || []);

    return {
      sunday: sundayResult,
      wednesday: wednesdayResult,
      special: specialResultArray,
      error: null,
      availableMonths: sortedAvailableMonths.reverse(),
      selectedMonth: actualSelectedMonth,
      staticWednesdayCentersList: wednesdayStaticCenters, // Pass this to the client if needed later
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
 * Helper function to pivot Sun/Wed data, calculate averages, apply custom sorting.
 * Ensures all centers from `allRelevantCenters` are included in the output.
 * Uses ymd (in SCRIPT_TIMEZONE) for internal logic, display values for cells.
 */
function processAndSortData(dataItems, selectedMonth, isSundayData = false, allRelevantCenters = null) {
  const scriptTimeZone = Session.getScriptTimeZone();
  
  // Initialize dataItems as an empty array if it's null or undefined, to simplify later logic
  dataItems = dataItems || [];
  allRelevantCenters = allRelevantCenters || [];


  let processedData = {}; // Holds data extracted from dataItems, keyed by centreName
  let uniqueDateInfo = {}; // Stores { "dd-MM-yyyy_display": "yyyy-MM-dd_sortable" }

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

  // Iterate through allRelevantCenters (customCentreOrder for Sunday, staticWednesdayCenters for Wednesday)
  // This ensures all these centers appear in the output, in the specified order.
  allRelevantCenters.forEach((centreName) => {
    const centreDataFromProcessed = processedData[centreName];
    let isMissingData = false;

    if (sortedDateKeys.length > 0) { // Only consider missing if there were actual dates for this event type in the month
      if (!centreDataFromProcessed) {
        // Centre has no data entries at all this month for this event type
        isMissingData = true;
      } else {
        // Check if this centre is missing data for any specific date
        for (const dateKey of sortedDateKeys) {
          if (!centreDataFromProcessed.dateData[dateKey]) {
            isMissingData = true; // Missing data for at least one event date
            break;
          }
        }
      }
    }
    // If sortedDateKeys.length is 0 (no events of this type in the data for this month),
    // then isMissingData remains false.

    const averageSangat =
      centreDataFromProcessed && centreDataFromProcessed.sangatCount > 0
        ? (
            centreDataFromProcessed.totalSangatSum /
            centreDataFromProcessed.sangatCount
          ).toFixed(1)
        : "N/A";

    finalTemplateData.push({
      centre: centreName,
      isBold: boldCentres.includes(centreName), // boldCentres applies to both Sun & Wed if names match
      dateData: centreDataFromProcessed ? centreDataFromProcessed.dateData : {},
      averageSangat: averageSangat,
      isMissingData: isMissingData,
    });
    centresAlreadyAddedToFinal.add(centreName);
  });

  // Add any remaining centres from processedData that were not in allRelevantCenters
  // These are centres that had data but weren't in the primary list.
  // For Wednesday, this block might not run if allRelevantCenters covers all possibilities.
  // For Sunday, this adds any centres with data not in customCentreOrder.
  Object.keys(processedData)
    .sort((a, b) => a.localeCompare(b)) // Alphabetical sort for these additional centers
    .forEach((centreName) => {
      if (!centresAlreadyAddedToFinal.has(centreName)) {
        const centreData = processedData[centreName];
        let isMissingDataForExtraCenter = false;
        if (sortedDateKeys.length > 0 && centreData) {
           for (const dateKey of sortedDateKeys) {
            if (!centreData.dateData[dateKey]) {
              isMissingDataForExtraCenter = true;
              break;
            }
          }
        } else if (sortedDateKeys.length > 0 && !centreData) {
          // This case should technically be caught by the previous loop if allRelevantCenters is comprehensive
          // or this center wasn't in processedData at all.
          // If it means a center name was in processedData keys but object is null/undefined
          isMissingDataForExtraCenter = true;
        }


        const averageSangat =
          centreData.sangatCount > 0
            ? (centreData.totalSangatSum / centreData.sangatCount).toFixed(1)
            : "N/A";
        finalTemplateData.push({
          centre: centreName,
          isBold: boldCentres.includes(centreName),
          dateData: centreData.dateData,
          averageSangat: averageSangat,
          // For centers not in the 'allRelevantCenters' list, 'isMissingData' logic might differ.
          // Typically, if they have *any* data, they are not "missing" in the same sense.
          // However, they could be missing specific dates if not all dates are filled.
          isMissingData: isMissingDataForExtraCenter,
        });
        centresAlreadyAddedToFinal.add(centreName); // Should not be strictly necessary here, but good practice
      }
    });

  // If after all processing, for a day type with allRelevantCenters defined,
  // and no data items were found at all (dataItems was empty),
  // but there are sortedDateKeys (which implies some data existed somewhere for this month for this day type,
  // though this scenario is unlikely if dataItems is empty for this specific call),
  // the isMissingData flag would have been set.
  // If dataItems is empty AND sortedDateKeys is empty, finalTemplateData will list allRelevantCenters
  // with N/A averages and isMissingData = false, which is correct.

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
 * Retrieves a list of static Wednesday center names from the "Configuration" sheet.
 * Assumes center names are in the first column (Column A), starting from row 1 or 2.
 * @return {Array<string>} An array of center names. Returns empty if sheet/data not found.
 */
function getStaticWednesdayCentersFromSheet() {
  const configSheetName = "Configuration";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(configSheetName);

  if (!sheet) {
    Logger.log(`Error: Sheet "${configSheetName}" not found. Cannot retrieve static Wednesday centers.`);
    return [];
  }

  try {
    // Get data from the first column. getLastRow might be too much if the column is sparsely populated
    // but other columns have more data. Range up to where data exists in column A.
    const lastRowWithDataInColA = sheet.getRange("A1:A").getValues().filter(String).length;
    if (lastRowWithDataInColA === 0) {
      Logger.log(`No data found in column A of "${configSheetName}" sheet.`);
      return [];
    }

    const range = sheet.getRange(1, 1, lastRowWithDataInColA); // Column A, down to the last row with content in A
    const values = range.getValues();
    const centers = values.map(row => row[0].toString().trim()).filter(name => name !== "");

    Logger.log(`Retrieved ${centers.length} static Wednesday centers from "${configSheetName}": ${centers.join(", ")}`);
    return centers;
  } catch (e) {
    Logger.log(`Error reading static Wednesday centers from "${configSheetName}": ${e.message} Stack: ${e.stack}`);
    return [];
  }
}


/**
 * Finds missing date entries for Sunday and Wednesday satsangs for a given month.
 * Returns an array of objects, each detailing the center, day, and missing dates.
 */
function findMissingDateEntries(selectedMonth) {
  const scriptTimeZone = Session.getScriptTimeZone();
  const currentMonth =
    selectedMonth ||
    Utilities.formatDate(new Date(), scriptTimeZone, "yyyy-MM");

  // Pass null for sundaySortPref as it's not relevant for missing date calculation here
  const processedDataResult = getDataForDashboard(currentMonth, null);

  if (processedDataResult.error) {
    Logger.log(`Error in findMissingDateEntries while fetching dashboard data: ${processedDataResult.error}`);
    // Depending on how you want to handle this, you might throw an error
    // or return an empty array or an array with an error object.
    // For now, let's throw, so the client-side gets a failure.
    throw new Error(processedDataResult.error);
  }

  const results = []; // Will store objects: { centerName: string, day: string, dates: string[] }

  // Check Sundays
  const sundayResult = processedDataResult.sunday;
  if (
    sundayResult &&
    sundayResult.templateData &&
    sundayResult.templateData.length > 0
  ) {
    const sundayDates = sundayResult.sortedDates || []; // These are 'dd-MM-yyyy'
    if (sundayDates.length > 0) { // Only proceed if there were Sundays this month
      sundayResult.templateData.forEach((center) => {
        // The `isMissingData` flag on `center` object (if populated by processAndSortData)
        // already tells us if this center is missing *any* Sunday.
        // However, to get the *specific* missing dates, we still need to check.
        const centerName = center.centre;
        const missingDates = sundayDates.filter((date) => !center.dateData[date]);
        if (missingDates.length > 0) {
          results.push({ centerName: centerName, day: "Sunday", dates: missingDates });
        }
      });
    }
  }

  // Check Wednesdays
  const wednesdayResult = processedDataResult.wednesday;
  // Get static list from Configuration sheet
  const staticWednesdayCenters = getStaticWednesdayCentersFromSheet();

  if (staticWednesdayCenters.length === 0) {
    Logger.log("findMissingDateEntries: No static Wednesday centers configured. Skipping Wednesday check.");
  } else if (
    wednesdayResult &&
    wednesdayResult.sortedDates &&
    wednesdayResult.sortedDates.length > 0
  ) {
    const wednesdayDates = wednesdayResult.sortedDates || []; // These are 'dd-MM-yyyy'

    staticWednesdayCenters.forEach(centerName => {
      // Find the center in the data returned by getDataForDashboard.
      // processAndSortData should now ensure all staticWednesdayCenters are in templateData.
      const centerData = (wednesdayResult.templateData || []).find(row => row.centre === centerName);
      let missingDatesForThisCenter = [];

      if (!centerData) {
        // This case should ideally not happen if processAndSortData correctly includes all static centers.
        // If it does, it means the center had no data entries at all.
        Logger.log(`findMissingDateEntries: Center '${centerName}' not found in Wednesday processed data. Assuming all dates missing.`);
        missingDatesForThisCenter = wednesdayDates.slice(); // All dates are missing
      } else {
        // Check which specific dates are missing for this center
        missingDatesForThisCenter = wednesdayDates.filter(date => !centerData.dateData[date]);
      }

      if (missingDatesForThisCenter.length > 0) {
        results.push({ centerName: centerName, day: "Wednesday", dates: missingDatesForThisCenter });
      }
    });
  } else if (staticWednesdayCenters.length > 0 && wednesdayResult && wednesdayResult.sortedDates && wednesdayResult.sortedDates.length === 0) {
    // Static centers exist, but no Wednesday satsangs were reported at all in the month.
    // This means all static centers are "missing" all potential Wednesdays (if any were expected).
    // However, without any `wednesdayResult.sortedDates`, there are no dates to report as missing.
    // This state is more "No Wednesday data this month" rather than specific missing entries.
    // So, we don't add anything to results here. The client can show "No data".
    Logger.log("findMissingDateEntries: Static Wednesday centers exist, but no Wednesday dates found in data for the month.");
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
