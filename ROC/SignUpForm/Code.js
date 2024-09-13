/**
 * Rural Outreach Clinic (ROC) Sign-Up Form and Scheduling Script
 * 
 * This script manages the sign-up process, match list generation, and email communications
 * for the Rural Outreach Clinic program. It interacts with Google Sheets and Forms to automate
 * the scheduling process for rural outreach clinics.
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

const DEBUG = true;

// Sheet IDs
const SHEETS_ID = {
  TRACKER: "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU",  // Match tracker
  DATES: "1eztpAJ1JSWps9SiHl6pHV_6y6pWZPZ_i9tP029nw7JQ",    // Clinic dates
  MATCH: "1FW1wiYIG9LvCEgFfpgf6c9GyKljNQbwEZkoRUE3lYWg",    // Match list
  SIGN: "1V4xGSO3RbAJIdsAvPdTU2raJ8xYyWBKuPHj9rql52S4",     // Form responses
  PEOPLE: "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"    // Contact info
};

const FORMS_ID = {
  OFFICIAL: "1FAIpQLSdeCfV1VzuK3clXnHaExcLQ88ekVoKWMMvTKql97WFr8p9WZQ",  // Main form
  MOD: "1FAIpQLSfPWCXGNPvNqVWOz2y0GYORi2_jxhtHU-vnvveXsZoUX8E54w",      // Modification form
};

const NAMES_ID = "2021179574";  // data-item-id for names list

// Signup timing (in days)
const SIGNUP_DAYS = {
  LEAD: 7,    // Open signup
  MANAGE: 3,  // Send preliminary match to manager
  CLOSE: 2    // Close signup
};

// Column indices
// Sign up sheet
const SIGN_INDEX = {
  NAME: 2,
  PTS_ALONE: 3,
  SPANISH: 4,
  SOC_POS: 5,
  ELECTIVE:6,
  FOLLOW: 7,
  CARPOOL: 8,
  COMMENTS: 9,
  DATE: 10,
  CLINIC_TYPE: 11
};

// Match list sheet
const MATCH_INDEX = {
  NAMES: 15,
  MANAGERS: "A8",
  TITLE: "A1",
  DATE: "A3",
  TIME: "C3",
  ADDRESS: "A5",
  CHALK: "D12",
  PHYS: "A12",
  INTERPRET: "C8",
  SHADOW: "D8",
  VOLUNT: "C12"
};

// NOTES:
//  Code infers year based on sheet order (MS1,2,3,4,PA1,2); could update but is already pretty simple



// *** ---------------------------------- *** // 

/**
 * Creates form submit and time-based triggers if they don't already exist.
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
 * Removes all triggers associated with the current project.
 */
function discontinueTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Updates the form date and manages the signup process.
 * This function is triggered daily to check upcoming clinics and manage the signup process.
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

    const typeCode = spreadsheet.getRange(`B${i}`).getValue();
    const clinicTypes = {
      Y: { type: "Yerington", address: "South Lyon Physicians Clinic, 213 S Whitacre St., Yerington,NV", rooms: 6 },
      SS: { type: "Silver Springs", address: "3595 Hwy 50, Suite 3 (In Lahontan Medical Complex), Silver Springs, NV", rooms: 5 },
      F: { type: "Fallon", address: "485 West B St Suite 101, Fallon, NV", rooms: 3 }
    };

    const clinicInfo = clinicTypes[typeCode] || { type: "Unknown", address: "Unknown", rooms: 0 };
    clinicInfo.time = clinicDate.getDay() === 0 ? "9am - 3pm" : clinicDate.getDay() === 6 ? "9am - 1pm" : "Unknown";
    clinicInfo.typeCode = typeCode;

    const links = {
      form: `https://docs.google.com/forms/d/e/${FORMS_ID.OFFICIAL}/viewform?usp=sf_link`,
      formMod: `https://docs.google.com/forms/d/e/${FORMS_ID.MOD}/viewform?usp=sf_link`,
      date: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.DATES}/edit?usp=sharing`,
      track: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`,
      match: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`
    };

    if (clinicDate.valueOf() === checkingDates.signUp.valueOf()) {
      handleSignUp(form, dateString, clinicDate, timeZone, clinicInfo, links);
    } else if (clinicDate.valueOf() === checkingDates.manage.valueOf()) {
      handlePreliminaryMatch(form, dateString, clinicDate, clinicInfo);
    } else if (clinicDate.valueOf() === checkingDates.close.valueOf()) {
      handleFinalMatch(dateString, clinicDate, clinicInfo);
    }
  }
}

/**
 * Updates the match list for a clinic based on sign-ups and participant scores.
 * This function handles the preliminary matching process, including:
 * - Gathering sign-ups for the current clinic date
 * - Calculating match scores for each participant
 * - Generating a sorted match list
 * - Updating the match list spreadsheet
 * - Sending a preliminary match list email to managers
 *
 * @param {Date} date - The date of the clinic
 * @param {string} type - The type code of the clinic (e.g., "Y" for Yerington)
 * @param {number} num_rooms - The number of available rooms for the clinic
 * @param {string} address - The address of the clinic
 */
function updateMatchList(date, type, num_rooms, address) {
  const clinicInfo = {
    time: date.getDay() === 0 ? "9AM - 3PM" : date.getDay() === 6 ? "9AM - 1PM" : "Unknown",
    title: {
      "Y": "Yerington ROC Clinic",
      "SS": "Silver Springs ROC Clinic",
      "F": "Fallon ROC Clinic"
    }[type] || "Unknown"
  };

  if (clinicInfo.time === "Unknown") {
    Logger.log("Issue with time extraction");
  }
  if (clinicInfo.title === "Unknown") {
    Logger.log("Problem with clinic type");
  }

  // Generate match list
  const matchList = generateMatchList(date, num_rooms);

  const copyMatchList = [...matchList]
  const actuallyMatched = setupMatchList(copyMatchList, clinicInfo, date, num_rooms, address);

  // Update match stats and gather sign-up information
  updateMatchStats(matchList, actuallyMatched, clinicInfo, date);
}

/**
 * Handles form submission for ROC clinic sign-ups.
 * 
 * This function is triggered when a form is submitted. It performs the following tasks:
 * 1. Logs the submitted name for error checking.
 * 2. Prevents resubmission for the same person and date.
 * 3. Updates the sign-up sheet with the new entry.
 * 4. Increments the sign-up counter in the tracker sheet for the participant.
 * 
 * @param {Object} e - The form submit event object containing response data.
 */
function onFormSubmit(e) {
  const form = FormApp.getActiveForm();
  const sheet = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];
  const sheetsTracker = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  
  const formResponse = e.response;
  const name = formResponse.getItemResponses()[0].getResponse();
  Logger.log(name); // Log name for error checking

  const [date, clinicTypeCode] = form.getDescription().split(";");
  const lastRow = sheet.getLastRow();
  
  // Prevent resubmission 
  const usedNames = sheet.getRange(2, SIGN_INDEX.NAME, lastRow - 1).getValues();
  const usedDates = sheet.getRange(2, SIGN_INDEX.DATE, lastRow - 1).getValues();
  
  for (let i = 0; i < lastRow - 2; i++) {
    if (name === usedNames[i][0] && new Date(date).valueOf() === usedDates[i][0].valueOf()) {
      Logger.log(`Form resubmission for ${name}`);
      return;
    }
  }

  if (isNaN(new Date(date).valueOf())) {
    Logger.log(`Bad date for ${name}: ${date}`);
    return;
  }

  // Update sign-up sheet
  sheet.getRange(lastRow, SIGN_INDEX.DATE).setValue(date);
  sheet.getRange(lastRow, SIGN_INDEX.CLINIC_TYPE).setValue(clinicTypeCode);

  // Update tracker sheet
  const nameArr = findCellByName(name);
  const trackerCell = sheetsTracker[nameArr[0]].getRange(nameArr[1], TRACK_INDEX.SIGNUPS);
  const currentSignups = trackerCell.getValue() || 0;
  
  if (!DEBUG) {
    trackerCell.setValue(currentSignups + 1);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Signups incremented from ${currentSignups} to ${currentSignups + 1}`);
  }
}