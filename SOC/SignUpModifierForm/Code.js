// Sheet and Form IDs
const SHEET_TRACKER = "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU"; // Match tracker sheet
const SHEET_SIGN = "1zCpaz2ketqM_EnYjnTTGPijs7_6ZKSu-VwtzcHkXA1g"; // Sheet associated with main form
const FORM_MAIN = "1vavuHv4ktebNNKUulTlUhY3JMjgYNv_6oZB0fCG_bms"; // ID of main form
const NAMES_ID = "2069885822"; // ID of the names list item in the form

// Sign up sheet column indices
const SIGN_NAME_NDX = 2;
const SIGN_PTS_ALONE_NDX = 3;
const SIGN_ELECTIVE_NDX = 4;
const SIGN_SOC_POS_NDX = 5;
const SIGN_DIET_NDX = 6;
const SIGN_COMMENTS_NDX = 7;
const SIGN_DATE_NDX = 8;
const SIGN_CLINIC_DATE_NDX = 9;

// Match tracker sheet column indices
const TRACK_LASTNAME_NDX = 1;
const TRACK_FIRSTNAME_NDX = 2;
const TRACK_SIGNUPS_NDX = 3;
const TRACK_MATCHES_NDX = 4;
const TRACK_NOSHOW_NDX = 5;
const TRACK_CXLLATE_NDX = 6;
const TRACK_CXLEARLY_NDX = 7;
const TRACK_DATE_NDX = 8;

// *** ---------------------------------- *** // 

/**
 * Creates form submit and time-based triggers.
 */
function createTriggers() {
  const form = FormApp.getActiveForm();
  const currentTriggers = ScriptApp.getProjectTriggers();

  // If triggers already exist, do not create new ones
  if (currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  // Create triggers
  ScriptApp.newTrigger("onFormSubmit").forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger("updateNames").timeBased().everyHours(1).create();
}

/**
 * Removes all triggers associated with the project.
 */
function discontinueTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Handles form submission events.
 * @param {Object} e - The event object containing form submission data.
 */
function onFormSubmit(e) {
  const form = FormApp.getActiveForm();
  const sheet = SpreadsheetApp.openById(SHEET_SIGN).getSheets()[0];
  
  const formResponse = e.response;
  Logger.log(formResponse.getItemResponses()[0].getResponse()); // Log name for error checking

  const date = form.getDescription().split(";")[0];
  const lastRow = sheet.getLastRow();
  
  const name = formResponse.getItemResponses()[0].getResponse();

  // Check for existing entries and toggle cancellation status
  const range = sheet.getRange(2, SIGN_NAME_NDX, lastRow - 1, 2);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    if (name === values[i][0] && new Date(date).valueOf() === values[i][1].valueOf()) {
      Logger.log(`Found sign up for ${name}`);
      const newName = name.slice(-3) !== "CXL" ? name + "CXL" : name.slice(0, -3);
      sheet.getRange(i + 2, SIGN_NAME_NDX).setValue(newName);
      break;
    }
  }

  updateNames();
}

/**
 * Updates the list of names in the form based on current sign-ups.
 */
function updateNames() {
  const form = FormApp.getActiveForm();
  const form_main = FormApp.openById(FORM_MAIN);
  const sheet_sign = SpreadsheetApp.openById(SHEET_SIGN).getSheets()[0];

  const descr = form_main.getDescription();
  const date = descr.split(";")[0];
  
  form.setDescription(descr);

  // Gather names of signups for current dated clinic
  const lastRow = sheet_sign.getLastRow();
  const range = sheet_sign.getRange(2, SIGN_DATE_NDX, lastRow - 1, 2);
  const values = range.getValues();

  const largeNameList = values
    .filter(row => new Date(date).valueOf() === row[0].valueOf())
    .map(row => row[1]);

  const namesList = form.getItemById(NAMES_ID).asListItem();

  // Set name options, defaulting to ["None"] if empty
  namesList.setChoiceValues(largeNameList.length > 0 ? largeNameList : ["None"]);
}