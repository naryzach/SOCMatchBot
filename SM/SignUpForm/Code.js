/**
 * Street Medicine (SM) Sign-Up Form and Scheduling Script
 * 
 * This script manages the sign-up process, match list generation, and email communications
 * for the Street Medicine program. It interacts with Google Sheets and Forms to automate
 * the scheduling process for street medicine clinics.
 * 
 * DEBUG Mode:
 * When DEBUG is set to true:
 * 1. All emails are sent to the Webmaster instead of their intended recipients.
 * 2. The TRACKER sheet is not modified. Actions that would modify the sheet are logged instead.
 * 3. Debug messages are logged to indicate when emails would be sent and sheets would be updated.
 * 
 * To run the script in normal mode, set DEBUG = false.
 * 
 * Important:
 * - Update SHEET ID's and FORM ID's with new instances as needed
 * - Run createTriggers() after updating the Match Tracker or changing form IDs
 * - Ensure clinic dates in the DATES sheet are up to date for proper automation flow
 */

// Constants for Google Sheets and Forms

const DEBUG = true;

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
      handleSignUp(form, dateString, clinicDate, timeZone, links);
    } else if (clinicDate.valueOf() === checkingDates.manage.valueOf()) {
      handlePreliminaryMatch(form, dateString, clinicDate, spreadsheet);
    } else if (clinicDate.valueOf() === checkingDates.close.valueOf()) {
      handleFinalMatch(dateString, clinicDate);
    }
  }
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
    
    if (!DEBUG) {
      const matchesCell = trackSheet.getRange(trackRow, TRACK_INDEX.MATCHES);
      matchesCell.setValue((matchesCell.getValue() || 0) + 1);
      trackSheet.getRange(trackRow, TRACK_INDEX.DATE).setValue(date);
    } else {
      Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Matches incremented, Date set to ${date}`);
    }

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
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("SMManager", "email"),
    subject: "Street Medicine Match List (Prelim) and Notes from Sign-ups",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: emailHtml,
    name: "SM Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Preliminary match list email sent to Webmaster instead of SM Manager for clinic on ${dateString}`);
  }

  FormApp.getActiveForm().deleteAllResponses();
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
  const clinicDate = new Date(date);
  if (isNaN(clinicDate.valueOf())) {
    Logger.log(`Invalid date: ${date} for ${name}`);
    return;
  }

  const usedNames = signUpSheet.getRange(2, SIGN_INDEX.NAME, lastRow - 1).getValues();
  const usedDates = signUpSheet.getRange(2, SIGN_INDEX.DATE, lastRow - 1).getValues();

  for (let i = 0; i < lastRow - 2; i++) {
    if (name === usedNames[i][0] && clinicDate.valueOf() === usedDates[i][0].valueOf()) {
      Logger.log(`Resubmission detected for ${name}`);
      return;
    }
  }

  // Set the date to the date of the clinic
  signUpSheet.getRange(lastRow + 1, SIGN_INDEX.DATE).setValue(date);

  // Update the sign-up counter in the tracker
  const nameArr = findCellByName(name);
  if (!nameArr) {
    Logger.log(`Could not find ${name} in tracker sheets`);
    return;
  }

  const [sheetIndex, rowIndex] = nameArr;
  const cell = trackerSheets[sheetIndex].getRange(rowIndex + 1, TRACK_INDEX.SIGNUPS);
  const currentValue = cell.getValue() || 0;
  
  if (!DEBUG) {
    cell.setValue(currentValue + 1);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Signups incremented from ${currentValue} to ${currentValue + 1}`);
  }
}