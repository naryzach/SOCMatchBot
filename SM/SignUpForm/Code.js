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
  MANAGERS: "A8",
  TITLE: "A1",
  DATE: "A3",
  TIME: "C3",
  ADDR: "B5",
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
  const dateSheet = spreadsheet.getSheets()[0];

  // Format date column
  const dateColumn = spreadsheet.getRange('A:A');
  dateColumn.setNumberFormat('dd-MM-yyyy');

  const today = new Date();
  const checkDates = {
    signUp: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.LEAD),
    close: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.CLOSE),
    manage: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.MANAGE)
  };

  const lastRow = dateSheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    const clinicDate = new Date(dateSheet.getRange(`A${i}`).getValue());
    const timeZone = `GMT-${clinicDate.getTimezoneOffset() / 60}`; // Note: This won't work east of Prime Meridian
    const dateString = Utilities.formatDate(clinicDate, timeZone, 'EEEE, MMMM dd, YYYY');

    const links = {
      form: `https://docs.google.com/forms/d/e/${FORMS_ID.OFFICIAL}/viewform?usp=sf_link`,
      formMod: `https://docs.google.com/forms/d/e/${FORMS_ID.MOD}/viewform?usp=sf_link`,
      date: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.DATES}/edit?usp=sharing`,
      track: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`,
      match: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`
    };

    if (clinicDate.valueOf() === checkDates.signUp.valueOf()) {
      handleSignUp(form, dateString, clinicDate, timeZone, links);
    } else if (clinicDate.valueOf() === checkDates.manage.valueOf()) {
      handlePreliminaryMatch(form, clinicDate, spreadsheet);
    } else if (clinicDate.valueOf() === checkDates.close.valueOf()) {
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
  const clinicTime = "8AM - 12PM";
  const clinicInfo = "Street Medicine Clinic";

  // generate match list
  const matchList = generateMatchList(date, numRooms);

  // setup match list
  const actuallyMatched = setupMatchList(matchList, clinicTime, clinicInfo, date, numRooms);

  // Update match stats and prepare manager email body
  updateMatchStats(actuallyMatched, date);
}

/**
 * Handles form submission event.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e - The form submit event object.
 */
function onFormSubmit(e) {
  const form = FormApp.getActiveForm();
  const signUpSheet = SpreadsheetApp.openById(form.getDestinationId()).getSheets()[0];
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
  const cell = trackerSheets[sheetIndex].getRange(rowIndex, TRACK_INDEX.SIGNUPS);
  const currentValue = cell.getValue() || 0;
  
  if (!DEBUG) {
    cell.setValue(currentValue + 1);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Signups incremented from ${currentValue} to ${currentValue + 1}`);
  }
}