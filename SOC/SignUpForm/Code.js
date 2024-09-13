/**
 * SOC (Student Outreach Clinic) Sign-Up Form and Scheduling Script
 * 
 * This script manages the sign-up process, match list generation, and email communications
 * for the Student Outreach Clinic. It interacts with Google Sheets and Forms to automate
 * the scheduling process.
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
 * - Update SHEET ID's with new instances (link new sheet to responses first)
 * - Run createTriggers() after Match Tracker is updated
 * - Update clinic dates for proper flow of automation
 */

const DEBUG = true;

// Sheet IDs
const SHEETS_ID = {
  TRACKER: "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU",  // Match tracker
  DATES: "1vhqF4JpN9HZwdqkMou_AQmC61mOA49mh57N2RIzpZHY",    // Clinic dates
  MATCH: "1hUJJqmnqrDD7e6n9MLtro6WIneTQl1o76atptyssig4",    // Match list
  SIGN: "1zCpaz2ketqM_EnYjnTTGPijs7_6ZKSu-VwtzcHkXA1g",     // Form responses
  PEOPLE: "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"    // Contact info
};

const FORMS_ID = {
  OFFICIAL: "1FAIpQLSf3361EGf494smoIhOwi8EVAgKhYR3IPqmm6SET-2RiHaLAdw",  // Main form
  MOD: "1FAIpQLSfwAJj6RyqB-u4QBVclWrFi5R9FokTEHS9jB1k8PidO75xtAQ",      // Modification form
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
  PTS_ALONE: 3,
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
  CHALK_TALK: "C10",
  LIAISON: "B8"
};

// Date sheet
const DATE_INDEX = {
  ROOMS: "C2",
  DEFAULT_NUM_ROOMS: "10"
};

// NOTES:
//  Code infers year based on sheet order (MS1,2,3,4,PA1,2); could update but is already pretty simple

// *** ---------------------------------- *** // 

/**
 * Creates installable triggers for the form submit and daily update functions.
 * This function should be run manually after updating the Match Tracker.
 */
function createTriggers() {
  // Get the form object.
  var form = FormApp.getActiveForm();

  // Check if triggers are already set 
  var currentTriggers = ScriptApp.getProjectTriggers();
  if (currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  // Create triggers
  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();
  
  ScriptApp.newTrigger("updateForm")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Update name list
  updateStudents();
}

/**
 * Removes all existing triggers for the project.
 * This function can be used to clean up or reset the project's triggers.
 */
function discontinueTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Updates the form title and description based on upcoming clinic dates.
 * This function is triggered daily to keep the form up-to-date.
 * 
 * It performs the following tasks:
 * 1. Updates the form title and description for upcoming clinics
 * 2. Sends email notifications for sign-ups and room availability
 * 3. Closes sign-ups and generates match lists when appropriate
 * 4. Schedules the sending of match lists
 */
function updateForm() {
  const form = FormApp.getActiveForm();
  const spreadsheet = SpreadsheetApp.openById(SHEETS_ID.DATES);
  const dateSheet = spreadsheet.getSheets()[0];

  // Format date column
  const dateColumn = dateSheet.getRange('A:A');
  dateColumn.setNumberFormat('dd-MM-yyyy');

  const today = new Date();
  const checkDates = {
    lead: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.LEAD),
    close: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.CLOSE),
    manage: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.MANAGE),
  };

  const lastRow = dateSheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    const clinicDate = new Date(dateSheet.getRange(`A${i}`).getValue());
    const clinicTypeCode = dateSheet.getRange(`B${i}`).getValue().toString();

    const timeMap = {
      2: "6pm - 10pm",
      6: "8am - 12pm"
    };

    const typeMap = {
      W: { name: "Women's", manager: "WomenManager" },
      GP: { name: "Gen/Peds", manager: "GenPedManager" },
      GD: { name: "Geri/Derm", manager: "GeriDermManager" }
    };

    const tz = `GMT-${clinicDate.getTimezoneOffset() / 60}`;
    const dateString = Utilities.formatDate(clinicDate, tz, 'EEEE, MMMM dd, YYYY');

    const clinicInfo = {
      date: clinicDate,
      dateString: dateString,
      time: timeMap[clinicDate.getDay()] || "Unknown",
      type: typeMap[clinicTypeCode]?.name || "Unknown",
      managerEmail: GET_INFO(typeMap[clinicTypeCode]?.manager || "Webmaster", "email"),
      typeCode: clinicTypeCode
    };

    if (clinicDate.valueOf() === checkDates.lead.valueOf()) {
      handleSignUp(form, clinicInfo);
    } else if (clinicDate.valueOf() === checkDates.manage.valueOf()) {
      handleManagers(spreadsheet, clinicInfo);
    } else if (clinicDate.valueOf() === checkDates.close.valueOf()) {
      handleFinalMatch(form, spreadsheet, clinicInfo);
    }
  }
}

/**
 * Updates the match list based on sign-ups and applies changes to the Sheets file.
 * 
 * @param {Date} date - The date of the clinic
 * @param {string} type - The type of clinic (e.g., "W" for Women's, "GP" for Gen/Peds)
 * @param {number} num_rooms - The number of available rooms for the clinic
 * 
 * This function performs the following tasks:
 * 1. Retrieves sign-up information from the sign-up sheet
 * 2. Calculates match scores for each signed-up student
 * 3. Generates a match list based on scores and available rooms
 * 4. Updates the match list sheet with the generated matches
 * 5. Sends an email to managers with sign-up notes and dietary restrictions
 */
function updateMatchList(date, type, numRooms) {
  // Clinic time and type
  const clinicTimes = {
    2: "6PM - 10PM",
    6: "8AM - 12PM"
  };
  const clinicTime = clinicTimes[date.getDay()] || "Unknown";
  
  if (!clinicTimes[date.getDay()]) {
    Logger.log("Issue with time extraction");
  }

  // Clinic type and manager
  const clinicTypes = {
    "W": {
      title: "Women's Clinic",
      managerType: "WomenManager"
    },
    "GP": {
      title: "General & Pediatric Clinic",
      managerType: "GenPedManager"
    },
    "GD": {
      title: "Geriatrics & Dermatology Clinic",
      managerType: "GeriDermManager"
    }
  };

  const clinicInfo = clinicTypes[type] || {
    title: "Unknown",
    managerType: "Webmaster"
  };

  if (!clinicTypes[type]) {
    Logger.log("Problem with clinic type");
  }

  // Generate match list
  const matchList = generateMatchList(date, numRooms);

  // Setup match list
  const copyMatchList = [...matchList]
  const actuallyMatched = setupMatchList(copyMatchList, clinicTime, clinicInfo, date, numRooms);
  
  // Update match stats
  updateMatchStats(matchList, actuallyMatched, clinicInfo, date);
}

/**
 * A function that is called by the form submit trigger.
 * The parameter e contains information submitted by the user.
 * 
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e - The form submit event object
 * 
 * This function performs the following tasks:
 * 1. Retrieves the submitted form response
 * 2. Checks for duplicate submissions
 * 3. Updates the sign-up counter for the submitted student
 */
function onFormSubmit(e) {
  // Get form, sheets, and response data
  const form = FormApp.getActiveForm();
  const signSheet = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];
  const trackerSheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const formResponse = e.response;
  
  Logger.log(formResponse.getItemResponses()[0].getResponse()); // Log name for error checking

  // Extract clinic info from form description
  const [date, clinicTypeCode] = form.getDescription().split(";");
  
  // Get submitted name and find corresponding row in tracker
  const name = formResponse.getItemResponses()[0].getResponse();
  const nameArr = findCellByName(name);

  // Check for duplicate submissions
  const lastRow = signSheet.getLastRow();
  const usedNames = signSheet.getRange(2, SIGN_INDEX.NAME, lastRow - 1, 1).getValues();
  const usedDates = signSheet.getRange(2, SIGN_INDEX.DATE, lastRow - 1, 1).getValues();

  for (let i = 0; i < lastRow - 2; i++) {
    if (name === usedNames[i][0] && new Date(date).valueOf() === usedDates[i][0].valueOf()) {
      Logger.log(`${name}: Form resubmission`);
      return;
    }
  }

  if (isNaN(new Date(date).valueOf())) {
    Logger.log(`${name}: ${date} - Bad date`);
    return;
  }

  // Update sign-up sheet
  signSheet.getRange(lastRow, SIGN_INDEX.DATE).setValue(date);
  signSheet.getRange(lastRow, SIGN_INDEX.CLINIC_TYPE).setValue(clinicTypeCode);

  // Update sign-up counter in tracker
  const trackerSheet = trackerSheets[nameArr[0]];
  const signupsCell = trackerSheet.getRange(nameArr[1], TRACK_INDEX.SIGNUPS);
  const currentSignups = signupsCell.getValue() || 0;

  if (!DEBUG) {
    signupsCell.setValue(currentSignups + 1);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Signups = ${currentSignups + 1}`);
  }
}

/**
 * Sends the current match list to the Webmaster for review.
 * 
 * This function performs the following tasks:
 * 1. Generates a PDF of the current match list
 * 2. Sends an email to the Webmaster with the PDF attached
 */
function sendWebmasterList() {
  //var raw_date = new Date(spreadsheet.getRange("X2").getValue().toString());
  var file = makeMatchPDF(new Date()); // make the PDF of the match list
  MailApp.sendEmail({
    to: GET_INFO("Webmaster", "email"),
    subject:  "Current Match List",
    replyTo: GET_INFO("Webmaster", "email"),
    body: "Attched is the current match list as a PDF",
    attachments: [file.getAs(MimeType.PDF)],
    name: "SOC Scheduling Assistant"
  });
}