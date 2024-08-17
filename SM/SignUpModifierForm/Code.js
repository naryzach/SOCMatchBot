// Sheet URLs
const SHEET_SIGN = "1mKUVnFeCzI8x2w83rifGX9IA9VFliNTbiDLEMpynPoI"; // sheet associated with main form
const FORM_MAIN = "1xcgvZ9eJDsPT_uuekd4o5XT2LWYhCS7sDhmhDj9IM5I" // ID of main form
const NAMES_ID = "2069885822";

// Sign up sheet indexing
const SIGN_NAME_NDX = 2
const SIGN_PTS_ALONE_NDX = 3
const SIGN_ELECTIVE_NDX = 4
const SIGN_SOC_POS_NDX = 5
const SIGN_DIET_NDX = 6
const SIGN_COMMENTS_NDX = 7
const SIGN_DATE_NDX = 8
const SIGN_CLINIC_DATE_NDX = 9

// Match tracker sheet indexing
const TRACK_LASTNAME_NDX = 1
const TRACK_FIRSTNAME_NDX = 2
const TRACK_SIGNUPS_NDX = 3
const TRACK_MATCHES_NDX = 4
const TRACK_NOSHOW_NDX = 5
const TRACK_CXLLATE_NDX = 6
const TRACK_CXLEARLY_NDX = 7
const TRACK_DATE_NDX = 8

// *** ---------------------------------- *** // 

// Create a form submit installable trigger
// using Apps Script.
function createTriggers() {
  // Get the form object.
  var form = FormApp.getActiveForm();

  // Check if triggers are already set 
  var currentTriggers = ScriptApp.getProjectTriggers();
  if(currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  // Create triggers
  ScriptApp.newTrigger("onFormSubmit").forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger("updateNames").timeBased().everyHours(1).create();
}

// Remove triggers
function discontinueTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// A function that is called by the form submit
// trigger. The parameter e contains information
// submitted by the user.
function onFormSubmit(e) {
  var form = FormApp.getActiveForm();
  var sheet = SpreadsheetApp.openById(SHEET_SIGN).getSheets()[0];
  
  // Get the response that was submitted.
  var formResponse = e.response;
  Logger.log(formResponse.getItemResponses()[0].getResponse()); // log name for error checking

  var descr = form.getDescription().split(";");
  var date = descr[0];
  let lastRow = sheet
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
  
  var itemResponses = formResponse.getItemResponses();
  var name = itemResponses[0].getResponse();

  // Prevent resubmission 
  var usedNames = sheet.getRange(2, SIGN_NAME_NDX, lastRow-1).getValues();
  var usedDates = sheet.getRange(2, SIGN_DATE_NDX, lastRow-1).getValues();
  for(var i = 0; i < lastRow-1; i++) {
    if (name == usedNames[i][0] && 
        new Date(date).valueOf() == usedDates[i][0].valueOf()) {
      Logger.log(name);
      Logger.log("Found sign up");
      if (name.slice(-3) != "CXL") {
        sheet.getRange(i + 2, SIGN_NAME_NDX).setValue(name + "CXL");
      } else {
        // Toggle cancelation
        sheet.getRange(i + 2, SIGN_NAME_NDX).setValue(name.slice(0,-3));
      }
    } 
  }
  updateNames()
}

// Build the list of names from the sheet
function updateNames () {
  var form = FormApp.getActiveForm();
  var form_main = FormApp.openById(FORM_MAIN);
  var sheet_sign = SpreadsheetApp.openById(SHEET_SIGN).getSheets()[0];

  var descr = form_main.getDescription();
  var date = descr.split(";")[0];
  var largeNameList = [];

  form.setDescription(descr);

  // Gather names of signups for current dated clinic
  let lastRow = sheet_sign
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
  var sign_dates = sheet_sign.getRange(2, SIGN_DATE_NDX, lastRow).getValues();
  var sign_names = sheet_sign.getRange(2, SIGN_NAME_NDX, lastRow).getValues();

  for(var i = 0; i < lastRow-1; i++) {
    if (new Date(date).valueOf()  == sign_dates[i][0].valueOf()) {
      largeNameList.push(sign_names[i][0]);
    }
  }

  var namesList = form.getItemById(NAMES_ID).asListItem();

  // Generate name options from Match Tracker
  if (largeNameList.length == 0) largeNameList = ["None"]
  namesList.setChoiceValues(largeNameList);
}