// Constants for Google Sheets and Forms

// Sheet and Form IDs
const SHEETS_ID = {
  TRACKER: "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU",  // Match tracker
  DATES: "1NKmLqbXjvEhfbYOoc82pwO4mls51Ih3Gc-Zn-OoMixI",    // Clinic dates
  MATCH: "1Be6ux1UynZ_s4toTyla5tU4IajixB9XFvRXnfbCHqJ4",    // Match list
  SIGN: "1mKUVnFeCzI8x2w83rifGX9IA9VFliNTbiDLEMpynPoI",     // Form responses
  PEOPLE: "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"    // Contact info
};

const FORMS_ID = {
  OFFICIAL: "1FAIpQLSf2EyVFnzzznQN2Y1DK_hVLlr51MV9DM0-V_Jk-XKtb3JT9RA",  // Main form
  MOD: "1FAIpQLSc5q0BHkHx9hyJ57bFUFML-aKYti1EncUpwHAGSJJe9E_SnhQ",      // Modification form
};

const NAMES_ID = "2021179574";  // data-item-id for names list

// Signup timing (in days)
const SIGNUP_DAYS = {
  LEAD: 5,    // Open signup
  MANAGE: 3,  // Send preliminary match to manager
  CLOSE: 2    // Close signup
};

// Column indices
// Sign up sheet
const SIGN_INDEX = {
  NAME: 2,
  TRANSPORT: 3,
  ELECTIVE: 4,
  SOC_POS: 5,
  DIET: 6,
  COMMENTS: 7,
  DATE: 8,
  CLINIC_TYPE: 9
};

// Match tracker sheet
const TRACK_INDEX = {
  LASTNAME: 1,
  FIRSTNAME: 2,
  SIGNUPS: 3,
  MATCHES: 4,
  NOSHOW: 5,
  CXLLATE: 6,
  CXLEARLY: 7,
  DATE: 8
};

// Match list sheet
const MATCH_INDEX = {
  NAMES: 13,
  MANAGERS: "A6",
  TITLE: "A1",
  DATE: "A3",
  TIME: "C3",
  PHYS1: "A10",
  PHYS2: "B10",
  CHALK_TALK: "C10"
};

// Date sheet
const DATE_INDEX = {
  ROOMS: "C2",
  DEFAULT_ROOMS: "10"
};

// People sheet
const PEOPLE_INDEX = {
  CEO: 2,
  COO: 3,
  WEBMASTER: 4,
  GEN_PED: 5,
  WOMEN: 6,
  GERI_DERM: 7,
  DIME: 8,
  LAY: 9,
  ROC: 10,
  SM: 11,
  CLASS: 12
};

// NOTES:
//  Code infers year based on sheet order (MS1,2,3,4,PA1,2); could update but is already pretty simple

// *** ---------------------------------- *** // 

/**
 * Creates form submit and daily update triggers for the Google Form.
 * This function sets up two triggers:
 * 1. A form submit trigger that calls onFormSubmit when the form is submitted.
 * 2. A time-based trigger that calls updateForm daily at 12:00 PM.
 * If triggers are already set, it logs a message and exits.
 * It also calls updateStudents to refresh the list of students in the form.
 */
function createTriggers() {
  const form = FormApp.getActiveForm();
  const currentTriggers = ScriptApp.getProjectTriggers();

  if (currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("updateForm")
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  updateStudents();
}

/**
 * Removes all existing triggers for the current project.
 * This function can be used to clean up or reset the project's triggers.
 */
function discontinueTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Updates the form title and sends emails based on upcoming clinic dates.
 * This function performs the following tasks:
 * 1. Retrieves clinic dates from a spreadsheet.
 * 2. Checks for upcoming clinics that require action (sign-up, management, or closure).
 * 3. Updates the form title and description for upcoming clinics.
 * 4. Sends appropriate emails for each stage of the clinic process.
 * It uses predefined lead times (SIGNUP_LEAD_DAYS, SIGNUP_CLOSE_DAYS, SIGNUP_MANAGE_DAYS) to determine when to take action.
 */
function updateForm() {
  const form = FormApp.getActiveForm();
  const spreadsheet = SpreadsheetApp.openById(SHEETS_ID.DATES);
  const dateColumn = spreadsheet.getRange('A:A');
  dateColumn.setNumberFormat('dd-MM-yyyy');

  const today = new Date();
  const checkingDates = {
    signUp: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.LEAD),
    close: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.CLOSE),
    manage: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.MANAGE)
  };

  const lastRow = spreadsheet.getSheets()[0]
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();

  for (let i = 2; i <= lastRow; i++) {
    const clinicDate = new Date(spreadsheet.getRange(`A${i}`).getValue());
    const timeZone = `GMT-${clinicDate.getTimezoneOffset() / 60}`; // Note: This won't work east of Prime Meridian
    const dateString = Utilities.formatDate(clinicDate, timeZone, 'EEEE, MMMM dd, YYYY');

    const links = {
      form: `https://docs.google.com/forms/d/e/${FORMS_ID.OFFICIAL}/viewform?usp=sf_link`,
      formMod: `https://docs.google.com/forms/d/e/${FORMS_ID.MOD}/viewform?usp=sf_link`,
      date: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.DATES}/edit?usp=sharing`,
      track: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`,
      match: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`
    };

    if (clinicDate.valueOf() === checkingDates.signUp.valueOf()) {
      sendSignUpEmail(form, dateString, clinicDate, timeZone, links);
    } else if (clinicDate.valueOf() === checkingDates.manage.valueOf()) {
      sendPreliminaryMatchList(form, dateString, clinicDate, spreadsheet);
    } else if (clinicDate.valueOf() === checkingDates.close.valueOf()) {
      sendFinalMatchList(dateString, clinicDate);
    }
  }
}

/**
 * Sends sign-up email and updates the form for an upcoming clinic.
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form to update.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {string} timeZone - Time zone string for date formatting.
 * @param {Object} links - Object containing various relevant URLs.
 */
function sendSignUpEmail(form, dateString, clinicDate, timeZone, links) {
  updateStudents();
  form.setTitle(`Street Medicine Clinic Sign Up -- ${dateString} from 8am - 12pm`);
  form.setDescription(Utilities.formatDate(clinicDate, timeZone, 'MM/dd/YYYY'));
  form.setAcceptingResponses(true);

  const formCloseDate = new Date();
  formCloseDate.setDate(formCloseDate.getDate() + (SIGNUP_DAYS.LEAD - SIGNUP_DAYS.MANAGE));

  const htmlBody = HtmlService.createTemplateFromFile('SignUpEmail');
  htmlBody.date = dateString;
  htmlBody.close_date = Utilities.formatDate(formCloseDate, timeZone, 'EEEE, MMMM dd, YYYY');
  htmlBody.link = links.form;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");

  MailApp.sendEmail({
    to: GET_INFO("ClassLists", "email"),
    subject: `Sign up for Street Medicine Clinic on ${dateString}`,
    replyTo: GET_INFO("SMManager", "email"),
    htmlBody: htmlBody.evaluate().getContent(),
    name: "Street Medicine Scheduling Assistant"
  });
}

/**
 * Sends preliminary match list and closes the sign-up form.
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form to update.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet containing clinic information.
 */
function sendPreliminaryMatchList(form, dateString, clinicDate, spreadsheet) {
  form.setTitle("Sign Ups Closed.");
  form.setDescription("Thank you for your interest. Please check back when another clinic is closer.");
  form.setAcceptingResponses(false);

  const numRooms = parseInt(spreadsheet.getRange(DATE_INDEX.ROOMS).getValue()) || DATE_INDEX.DEFAULT_ROOMS;
  updateMatchList(clinicDate, numRooms);
}

/**
 * Sends final match list with PDF attachment to participants.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 */
function sendFinalMatchList(dateString, clinicDate) {
  const file = makeMatchPDF(clinicDate);
  const htmlBody = HtmlService.createTemplateFromFile('MatchEmail');
  htmlBody.date = dateString;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");

  MailApp.sendEmail({
    to: GET_INFO("ClassLists", "email"),
    subject: `Match list for Street Medicine Clinic on ${dateString}`,
    replyTo: GET_INFO("SMManager", "email"),
    htmlBody: htmlBody.evaluate().getContent(),
    attachments: [file.getAs(MimeType.PDF)],
    name: "Street Medicine Scheduling Assistant"
  });
}

/**
 * Updates the match list for a Street Medicine clinic and sends a preliminary email to managers.
 * 
 * This function performs the following tasks:
 * 1. Retrieves sign-up information from the sign-up sheet.
 * 2. Calculates match scores for each participant based on various factors:
 *    - Number of sign-ups and matches
 *    - SOC position and seniority
 *    - Last match date
 *    - Cancellation history
 * 3. Generates a sorted match list based on calculated scores.
 * 4. Updates the match list sheet with the selected participants.
 * 5. Updates match statistics in the tracker sheets.
 * 6. Prepares and sends an email to managers with the preliminary match list and sign-up notes.
 * 7. Clears all responses from the active form.
 * 
 * @param {Date} date - The date of the clinic.
 * @param {number} numRooms - The number of available rooms for the clinic.
 */
function updateMatchList(date, numRooms) {
  const sheetMatch = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  // Gather names of signups for current dated clinic
  const lastRow = sheetSign.getLastRow();
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow - 1, 1).getValues();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow - 1, 1).getValues();
  
  const largeNameList = signDates.reduce((acc, date, index) => {
    if (date[0].valueOf() === date.valueOf()) {
      acc.push(signNames[index][0]);
    }
    return acc;
  }, []);

  const namesWithScores = {};

  // Generate match list
  largeNameList.forEach(name => {
    const nameRowIndex = signNames.findIndex(row => row[0] === name) + 2;
    const nameArr = findCellByName(name);
    
    // Check for errors reading names
    if (nameArr[0] === -1) {
      Logger.log(`Name error: ${name}`);
      if (name.slice(-3) === "CXL") {
        const newNameArr = findCellByName(name.slice(0, -3));
        if (newNameArr[0] === -1) return;

        // Update the sign up counter if cancellation
        const cxlEarlyCell = sheetsTrack[newNameArr[0]].getRange(newNameArr[1] + 1, TRACK_INDEX.CXLEARLY);
        cxlEarlyCell.setValue((cxlEarlyCell.getValue() || 0) + 1);
      }
      return;
    }
    
    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1] + 1;
    const signUps = trackSheet.getRange(trackRow, TRACK_INDEX.SIGNUPS).getValue() || 0;
    const matches = trackSheet.getRange(trackRow, TRACK_INDEX.MATCHES).getValue() || 0;
    const noShow = trackSheet.getRange(trackRow, TRACK_INDEX.NOSHOW).getValue() || 0;
    const cxlLate = trackSheet.getRange(trackRow, TRACK_INDEX.CXLLATE).getValue() || 0;
    const cxlEarly = trackSheet.getRange(trackRow, TRACK_INDEX.CXLEARLY).getValue() || 0;
    const lastDate = trackSheet.getRange(trackRow, TRACK_INDEX.DATE).getValue();
    const socPos = sheetSign.getRange(nameRowIndex, SIGN_INDEX.SOC_POS).getValue();

    // Calculate match score
    let matchScore = signUps - matches;

    // Elective and SOC position additions
    if (socPos === "Yes" && (nameArr[0] === 0 || nameArr[0] === 1)) {
      matchScore *= 2; // Only slightly bias SOC members rather than rank in a hierarchy
    }

    // Add points based on seniority
    const seniorityPoints = [5000, 50, 500, 1000, 0, 0];
    matchScore += seniorityPoints[nameArr[0]] || 0;

    // Never been matched addition or add fractional points based on last match
    if (lastDate === "") {
      matchScore += 25;
    } else {
      const daysSince = (new Date() - new Date(lastDate)) / (1000 * 60 * 60 * 24);
      matchScore += daysSince / 365;
    }

    // Cancellation penalty
    matchScore -= (noShow * 3) + (cxlLate * 2) + cxlEarly;

    namesWithScores[name] = matchScore;
  });

  Logger.log(namesWithScores);

  // Generate match list based on points
  const sortedNames = Object.entries(namesWithScores)
    .sort((a, b) => b[1] - a[1])
    .map(entry => entry[0]);

  const matchList = sortedNames.slice(0, Math.min(sortedNames.length, numRooms * 2));

  Logger.log(matchList);

  // Clear and update Match List Sheet
  const clearRange = sheetMatch.getRange(MATCH_INDEX.NAMES, 1, 25, 3);
  clearRange.clearContent().setBorder(false, false, false, false, false, false);

  sheetMatch.getRange(MATCH_INDEX.DATE).setValue(date);
  sheetMatch.getRange(MATCH_INDEX.TIME).setValue("8AM - 12PM");

  const actuallyMatched = [];
  const numSlots = Math.min(matchList.length, numRooms);

  // Fill rooms with people who can see patients alone
  for (let i = 0; i < numSlots; i++) {
    const name = matchList[i];
    actuallyMatched.push(name);
    const nameArr = findCellByName(name);
    const firstName = sheetsTrack[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
    const lastName = sheetsTrack[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();
    
    const matchRow = sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3);
    matchRow.setValues([
      [`Student ${i + 1}`, 
       `${firstName} ${lastName}, ${getYearTag(nameArr[0])}`, 
       "_____________________________________________\n_____________________________________________\n_____________________________________________"]
    ]).setBorder(true, true, true, true, true, true);
  }

  // Update match stats and prepare manager email body
  let managerEmailBody = "";
  actuallyMatched.forEach(name => {
    const nameArr = findCellByName(name);
    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1] + 1;
    
    const matchesCell = trackSheet.getRange(trackRow, TRACK_INDEX.MATCHES);
    matchesCell.setValue((matchesCell.getValue() || 0) + 1);
    trackSheet.getRange(trackRow, TRACK_INDEX.DATE).setValue(date);

    const nameRowIndex = signNames.findIndex(row => row[0] === name) + 2;
    const transport = sheetSign.getRange(nameRowIndex, SIGN_INDEX.ELECTIVE).getValue();
    const comments = sheetSign.getRange(nameRowIndex, SIGN_INDEX.COMMENTS).getValue();

    managerEmailBody += `${name} -- Reliable transport: ${transport}; Comments: ${comments}\n`;
  });

  // Send email with the preliminary match list for Managers to update
  const htmlBody = HtmlService.createTemplateFromFile('PrelimMatchEmail');
  const timeZone = "GMT-" + String(date.getTimezoneOffset() / 60); // will not work east of Prime Meridian
  const dateString = Utilities.formatDate(date, timeZone, 'EEEE, MMMM dd, YYYY');
  const linkMatch = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`;
  const linkTrack = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`;
  
  htmlBody.date = dateString;
  htmlBody.link = linkMatch;
  htmlBody.link_track = linkTrack;
  htmlBody.sign_up_notes = managerEmailBody;
  
  const emailHtml = htmlBody.evaluate().getContent();
  MailApp.sendEmail({
    to: GET_INFO("SMManager", "email"),
    subject: "Street Medicine Match List (Prelim) and Notes from Sign-ups",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: emailHtml,
    name: "SM Scheduling Assistant"
  });

  FormApp.getActiveForm().deleteAllResponses();
}

/**
 * Creates a PDF of the match list for a given date.
 * @param {Date} date - The date of the clinic.
 * @return {GoogleAppsScript.Drive.File} The created PDF file.
 */
function makeMatchPDF(date) {
  // Format the PDF name
  const pdfName = `MatchList_${date.toISOString().split('T')[0]}.pdf`;
  
  // Get the match sheet
  const sheet = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];

  // Define the range for the PDF
  const firstRow = 0, firstCol = 0, lastCol = 4, lastRow = 30;

  // Construct the URL for PDF export
  const url = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/export` +
    '?format=pdf&' +
    'size=7&' +
    'fzr=true&' +
    'portrait=true&' +
    'fitw=true&' +
    'gridlines=false&' +
    'printtitle=false&' +
    'top_margin=0.25&' +
    'bottom_margin=0.25&' +
    'left_margin=0.25&' +
    'right_margin=0.25&' +
    'sheetnames=false&' +
    'pagenum=UNDEFINED&' +
    'attachment=true&' +
    `gid=${sheet.getSheetId()}&` +
    `r1=${firstRow}&c1=${firstCol}&r2=${lastRow}&c2=${lastCol}`;

  // Set up parameters for the URL fetch
  const params = { 
    method: "GET", 
    headers: { "authorization": `Bearer ${ScriptApp.getOAuthToken()}` } 
  };

  // Fetch the PDF as a blob
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(`${pdfName}.pdf`);

  // Get the folder where PDFs are stored
  const folder = DriveApp.getFoldersByName("MatchListsSM").next();

  // Create the PDF file in the folder
  folder.createFile(blob);

  // Return the created file
  return DriveApp.getFilesByName(`${pdfName}.pdf`).next();
}

/**
 * Handles form submission event.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e - The form submit event object.
 */
function onFormSubmit(e) {
  const form = FormApp.getActiveForm();
  const signUpSheet = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];
  const trackerSheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  
  const formResponse = e.response;
  const itemResponses = formResponse.getItemResponses();
  const name = itemResponses[0].getResponse();
  const date = form.getDescription();

  Logger.log(name); // Log name for error checking

  // Get the last row with data in the sign-up sheet
  const lastRow = signUpSheet.getLastRow();
  
  // Check for resubmission or invalid date
  if (isResubmissionOrInvalidDate(signUpSheet, name, date, lastRow)) {
    return;
  }

  // Set the date to the date of the clinic
  signUpSheet.getRange(lastRow + 1, SIGN_INDEX.DATE).setValue(date);

  // Update the sign-up counter in the tracker
  updateSignUpCounter(trackerSheets, name);
}

/**
 * Checks if the submission is a resubmission or has an invalid date.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sign-up sheet.
 * @param {string} name - The name of the submitter.
 * @param {string} date - The clinic date.
 * @param {number} lastRow - The last row with data in the sheet.
 * @return {boolean} True if resubmission or invalid date, false otherwise.
 */
function isResubmissionOrInvalidDate(sheet, name, date, lastRow) {
  const usedNames = sheet.getRange(2, SIGN_INDEX.NAME, lastRow - 1).getValues();
  const usedDates = sheet.getRange(2, SIGN_INDEX.DATE, lastRow - 1).getValues();
  const clinicDate = new Date(date);

  if (isNaN(clinicDate.valueOf())) {
    Logger.log(`Invalid date: ${date} for ${name}`);
    return true;
  }

  for (let i = 0; i < lastRow - 2; i++) {
    if (name === usedNames[i][0] && clinicDate.valueOf() === usedDates[i][0].valueOf()) {
      Logger.log(`Resubmission detected for ${name}`);
      return true;
    }
  }

  return false;
}

/**
 * Updates the sign-up counter for the given name in the tracker sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} sheets - The tracker sheets.
 * @param {string} name - The name of the person to update.
 */
function updateSignUpCounter(sheets, name) {
  const nameArr = findCellByName(name);
  if (!nameArr) {
    Logger.log(`Could not find ${name} in tracker sheets`);
    return;
  }

  const [sheetIndex, rowIndex] = nameArr;
  const cell = sheets[sheetIndex].getRange(rowIndex + 1, TRACK_INDEX.SIGNUPS);
  const currentValue = cell.getValue() || 0;
  cell.setValue(currentValue + 1);
}

/**
 * Builds and returns a sorted list of student names from the tracker sheets.
 * @return {string[]} An array of formatted student names.
 */
function buildNameList() {
  const sheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const studentNames = new Set(); // Use a Set to automatically handle duplicates

  sheets.forEach((sheet, sheetIndex) => {
    const yearTag = getYearTag(sheetIndex);
    if (!yearTag) return;

    const lastRow = sheet.getLastRow();
    const namesRange = sheet.getRange(2, TRACK_INDEX.LASTNAME, lastRow - 1, 2);
    const namesValues = namesRange.getValues();

    namesValues.forEach(([lastName, firstName]) => {
      if (lastName) {
        const formattedName = `${lastName}, ${firstName} (${yearTag})`;
        if (studentNames.has(formattedName)) {
          Logger.log(`Duplicate: ${formattedName}`);
        } else {
          studentNames.add(formattedName);
        }
      }
    });
  });

  const sortedNames = Array.from(studentNames).sort();
  Logger.log(sortedNames);
  return sortedNames;
}

/**
 * Finds a student's sheet index and row index given their formatted name.
 * @param {string} name - The formatted name of the student (e.g., "Doe, John (MS2)").
 * @returns {number[]|null} An array containing [sheetIndex, rowIndex], or null if not found.
 */
function findCellByName(name) {
  const sheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const [lastName, firstName] = name.slice(0, -6).split(", ");

  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const sheet = sheets[sheetIndex];
    const lastNames = sheet.getRange(2, TRACK_INDEX.LASTNAME, sheet.getLastRow() - 1, 1).getValues();
    const firstNames = sheet.getRange(2, TRACK_INDEX.FIRSTNAME, sheet.getLastRow() - 1, 1).getValues();

    for (let rowIndex = 0; rowIndex < lastNames.length; rowIndex++) {
      if (firstNames[rowIndex][0] === firstName && lastNames[rowIndex][0] === lastName) {
        return [sheetIndex, rowIndex + 1]; // +1 because we started from row 2
      }
    }
  }

  Logger.log(`Did not find name: ${name}`);
  return null;
}

/**
 * Updates the name list for the Google Form.
 */
function updateStudents() {
  const form = FormApp.getActiveForm();
  const namesList = form.getItemById(NAMES_ID).asListItem();
  const studentNames = buildNameList();
  namesList.setChoiceValues(studentNames);
}

/**
 * Returns the year tag based on the sheet index.
 * @param {number} sheetIndex - The index of the sheet.
 * @returns {string} The year tag (e.g., "MS1", "PA2"), or an empty string if invalid.
 */
function getYearTag(sheetIndex) {
  const tags = ["MS1", "MS2", "MS3", "MS4", "PA1", "PA2"];
  return sheetIndex < tags.length ? tags[sheetIndex] : "";
}

/**
 * Retrieves information about a person based on their position.
 * @param {string} position - The position of the person (e.g., "CEO", "COO", "Webmaster", etc.).
 * @param {string} info - The type of information to retrieve ("name" or "email").
 * @returns {string} The requested information or an error message if not found.
 */
function GET_INFO(position, info) {
  const sheet = SpreadsheetApp.openById(SHEETS_ID.PEOPLE).getSheets()[0];
  const positionMap = {
    CEO: PEOPLE_INDEX.CEO,
    COO: PEOPLE_INDEX.COO,
    Webmaster: PEOPLE_INDEX.WEBMASTER,
    GenPedManager: PEOPLE_INDEX.GEN_PED,
    WomenManager: PEOPLE_INDEX.WOMEN,
    GeriDermManager: PEOPLE_INDEX.GERI_DERM,
    DIMEManager: PEOPLE_INDEX.DIME,
    ROCManager: PEOPLE_INDEX.ROC,
    SMManager: PEOPLE_INDEX.SM,
    LayCouns: PEOPLE_INDEX.LAY,
    ClassLists: PEOPLE_INDEX.CLASS
  };

  // Check if the position exists in our map
  if (!(position in positionMap)) {
    return "Position Not Found";
  }

  const row = positionMap[position];
  const name = sheet.getRange(row, 2).getValue();
  const email = sheet.getRange(row, 3).getValue();

  // Return the requested information
  switch (info.toLowerCase()) {
    case "email":
      return email;
    case "name":
      return name;
    default:
      return "Bad Lookup";
  }
}
