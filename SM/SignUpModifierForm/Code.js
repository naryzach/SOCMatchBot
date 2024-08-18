// Sheet and Form IDs
const SHEET_SIGN_ID = "1mKUVnFeCzI8x2w83rifGX9IA9VFliNTbiDLEMpynPoI"; // Sheet associated with main form
const FORM_MAIN_ID = "1xcgvZ9eJDsPT_uuekd4o5XT2LWYhCS7sDhmhDj9IM5I"; // ID of main form
const NAMES_ITEM_ID = "2069885822"; // ID of the names list item in the form

// Sign up sheet column indices
const SIGN_INDEX = {
  NAME: 2,
  PTS_ALONE: 3,
  ELECTIVE: 4,
  SOC_POS: 5,
  DIET: 6,
  COMMENTS: 7,
  DATE: 8,
  CLINIC_DATE: 9
};

// Match tracker sheet column indices
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

// *** ---------------------------------- *** // 

/**
 * Creates form submit and time-based triggers for the active form.
 * Checks if triggers are already set before creating new ones.
 */
function createTriggers() {
  const form = FormApp.getActiveForm();
  const currentTriggers = ScriptApp.getProjectTriggers();

  if (currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  ScriptApp.newTrigger("onFormSubmit").forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger("updateNames").timeBased().everyHours(1).create();
}

/**
 * Removes all existing triggers for the current project.
 */
function discontinueTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Handles form submission events.
 * Prevents resubmission and toggles cancellation status.
 * @param {Object} e - The form submit event object.
 */
function onFormSubmit(e) {
  const form = FormApp.getActiveForm();
  const sheet = SpreadsheetApp.openById(SHEET_SIGN_ID).getSheets()[0];
  const formResponse = e.response;
  const name = formResponse.getItemResponses()[0].getResponse();
  
  Logger.log(name); // Log name for error checking

  const date = form.getDescription().split(";")[0];
  const lastRow = sheet.getLastRow();
  
  const usedNames = sheet.getRange(2, SIGN_INDEX.NAME, lastRow - 1, 1).getValues().flat();
  const usedDates = sheet.getRange(2, SIGN_INDEX.DATE, lastRow - 1, 1).getValues().flat();

  for (let i = 0; i < lastRow - 1; i++) {
    if (name === usedNames[i] && new Date(date).valueOf() === usedDates[i].valueOf()) {
      Logger.log("Found sign up");
      const cell = sheet.getRange(i + 2, SIGN_INDEX.NAME);
      cell.setValue(cell.getValue().endsWith("CXL") ? name.slice(0, -3) : name + "CXL");
      break;
    }
  }

  updateNames();
}

/**
 * Updates the list of names in the form based on the sign-up sheet.
 * Filters names for the current clinic date.
 */
function updateNames() {
  const form = FormApp.getActiveForm();
  const formMain = FormApp.openById(FORM_MAIN_ID);
  const sheetSign = SpreadsheetApp.openById(SHEET_SIGN_ID).getSheets()[0];

  const date = formMain.getDescription().split(";")[0];
  form.setDescription(formMain.getDescription());

  const lastRow = sheetSign.getLastRow();
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow - 1, 1).getValues().flat();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow - 1, 1).getValues().flat();

  const currentDateValue = new Date(date).valueOf();
  const largeNameList = signDates.reduce((acc, dateValue, index) => {
    if (new Date(dateValue).valueOf() === currentDateValue) {
      acc.push(signNames[index]);
    }
    return acc;
  }, []);

  const namesList = form.getItemById(NAMES_ITEM_ID).asListItem();
  namesList.setChoiceValues(largeNameList.length > 0 ? largeNameList : ["None"]);
}