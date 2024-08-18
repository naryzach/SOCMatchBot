// Sheet URLs
const SHEET_SIGN_ID = "1FNhu-fzahjrlL0K--Mh-gSxBmDqqYWrBPG1VBOjbttQ"; // sheet associated with main form
const FORM_MAIN_ID = "1Xpw6O7zK0_9z_yDCzqIz9U_T-Zcv11psDuI20S0Gc0c" // ID of main form
const NAMES_ITEM_ID = "2069885822";

// Sign up sheet column indices
const SIGN_INDEX = {
  NAME: 2,
  PTS_ALONE: 3,
  SPANISH: 4,
  SOC_POS: 5,
  WATCHER: 6,
  TRANSPORT: 7,
  COMMENTS: 8,
  DATE: 9,
  CLINIC_TYPE: 10
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
 * Updates the list of names in the form based on current sign-ups.
 */
function updateNames() {
  const form = FormApp.getActiveForm();
  const form_main = FormApp.openById(FORM_MAIN_ID);
  const sheet_sign = SpreadsheetApp.openById(SHEET_SIGN_ID).getSheets()[0];

  const descr = form_main.getDescription();
  const date = descr.split(";")[0];
  
  // Gather names of signups for current dated clinic
  const lastRow = sheet_sign.getLastRow();
  const range = sheet_sign.getRange(2, SIGN_INDEX.DATE, lastRow - 1, 2);
  const values = range.getValues();

  const largeNameList = values
    .filter(row => new Date(date).valueOf() === row[0].valueOf())
    .map(row => row[1]);

  const namesList = form.getItemById(NAMES_ITEM_ID).asListItem();

  // Set name options, defaulting to ["None"] if empty
  namesList.setChoiceValues(largeNameList.length > 0 ? largeNameList : ["None"]);
}