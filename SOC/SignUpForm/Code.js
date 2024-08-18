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
  CHALK_TALK: "C10"
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
  ScriptApp.newTrigger("onFormSubmit").forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger("updateForm").timeBased().atHour(8).everyDays(1).create();

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
  dateSheet.getRange('A:A').setNumberFormat('dd-MM-yyyy');

  const today = new Date();
  const checkDates = {
    lead: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.LEAD),
    close: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.CLOSE),
    manage: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.MANAGE),
    ceo: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1)
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
function updateMatchList(date, type, num_rooms) {
  var sheet_match = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];
  var sheets_track = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  var sheet_sign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  var largeNameList = [];

  // Gather names of signups for current dated clinic
  let lastRow = sheet_sign
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
  var sign_dates = sheet_sign.getRange(2, SIGN_INDEX.DATE, lastRow).getValues();
  var sign_names = sheet_sign.getRange(2, SIGN_INDEX.NAME, lastRow).getValues();
  for(var i = 0; i < lastRow-1; i++) {
    if (date.valueOf() == sign_dates[i][0].valueOf()) {
      largeNameList.push(sign_names[i][0]);
    }
  }

  var nameArr = [];
  var matchScore = 0;
  var matches = 0;
  var signUps = 0;
  var noShow = 0;
  var cxlEarly = 0;
  var cxlLate = 0;
  var lastDate = "";
  var ptsAlone = "No";
  var fourthYrElect = "No";
  var socPos = "No";
  var namesWithScores = {};

  // Generate match list
  for (name in largeNameList) {
    matchScore = 0; 

    // Find which row the name is found on the sign up sheet
    var name_row_ndx = 0;
    for(var i = 0; i < lastRow-1; i++) {
      if (date.valueOf() == sign_dates[i][0].valueOf()) {
        if (sign_names[i][0] == largeNameList[name]) {
          name_row_ndx = i + 2; // List index offset from sheet
        }
      }
    }

    // Grab data on all sign ups
    nameArr = findCellByName(largeNameList[name])
    
    // Check for errors reading names
    if (nameArr[0] == -1) {
      Logger.log("Name error");
      Logger.log(largeNameList[name]);
      if (largeNameList[name].slice(-3) == "CXL") {
        // Grab new name w/o "CXL"
        nameArr = findCellByName(largeNameList[name].slice(0,-3))
        if (nameArr[0] == -1) continue;

        // Update the sign up counter if cancellation
        var tmp = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.CXLEARLY).getValue();
        if (tmp == "") { 
          tmp = 0;
        }
        if (!DEBUG) {
          sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.CXLEARLY).setValue(parseInt(tmp) + 1);
        } else {
          Logger.log(`DEBUG: Would update TRACKER sheet for ${largeNameList[name].slice(0,-3)}: CXLEARLY = ${parseInt(tmp) + 1}`);
        }
      }

      // Do not try to match that name
      continue;
    }
    signUps = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.SIGNUPS).getValue());
    matches = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.MATCHES).getValue());
    noShow = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.NOSHOW).getValue());
    cxlLate = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.CXLLATE).getValue());
    cxlEarly = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.CXLEARLY).getValue());
    lastDate = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.DATE).getValue();

    // Grab form submission data
    fourthYrElect = sheet_sign.getRange(name_row_ndx, SIGN_INDEX.ELECTIVE).getValue();
    socPos = sheet_sign.getRange(name_row_ndx, SIGN_INDEX.SOC_POS).getValue();

    // Calculate base match score
    matchScore = signUps - matches;

    // Adjust score based on student status and position
    if (socPos == "Yes" && nameArr[0] <= 1) { // SOC members (MS1/2s)
      matchScore *= 2;
    }
    if (fourthYrElect == "Yes" && nameArr[0] == 3) { // MS4s on elective
      matchScore += 500;
    }

    // Add points based on seniority
    const seniorityPoints = [0, 50, 500, 1000, 0, 0];
    matchScore += seniorityPoints[nameArr[0]] || 0;

    // Adjust for last match date
    if (lastDate == "") {
      matchScore += 25; // Never been matched
    } else {
      const daysSince = (new Date() - new Date(lastDate)) / (1000 * 60 * 60 * 24);
      matchScore += daysSince / 365;
    }

    // Apply cancellation penalties
    matchScore -= (noShow * 3 + cxlLate * 2 + cxlEarly) || 0;

    // Create dictionary of name (key) and score (value)
    namesWithScores[largeNameList[name]] = matchScore;
  }

  Logger.log("Match scores");
  Logger.log(namesWithScores);

  // Generate match list based on points
  var matchList = [];
  var sorted = Object.keys(namesWithScores).map(function(key) {
    return [key, namesWithScores[key]];
  });
  sorted.sort(function(first, second) {
    return second[1] - first[1];
  });
  // Need to check date for last match
  for (i = 0; i < (sorted.length < (num_rooms * 2) ? sorted.length : (num_rooms * 2)); i++) {
    matchList.push(sorted[i][0]);
  }

  Logger.log("Prelim match list");
  Logger.log(matchList);
  //GmailApp.sendEmail(ERROR_EMAIL, "Match List", matchList);

  // Clear Match List Sheet file names
  for (i = 0; i < 25; i++) { // 25 is an arbitrary choice. Should be more than max possible
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 1).setValue("");
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue("");
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 3).setValue("");
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(false, false, false, false, false, false);
  }
  // Clear physicians and chalk talk
  sheet_match.getRange(MATCH_INDEX.PHYS1).setValue("");
  sheet_match.getRange(MATCH_INDEX.PHYS2).setValue("");
  sheet_match.getRange(MATCH_INDEX.CHALK_TALK).setValue("");

  // Update Match List Sheet header
  const clinicTimes = {
    2: "6PM - 10PM",
    6: "8AM - 12PM"
  };
  const clinic_time = clinicTimes[date.getDay()] || "Unknown";
  
  if (!clinicTimes[date.getDay()]) {
    Logger.log("Issue with time extraction");
  }

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

  const clinic_title = clinicInfo.title;
  const manager_names = GET_INFO(clinicInfo.managerType, "name");
  const manager_emails = GET_INFO(clinicInfo.managerType, "email");

  sheet_match.getRange(MATCH_INDEX.TITLE).setValue(clinic_title);
  sheet_match.getRange(MATCH_INDEX.DATE).setValue(date);
  sheet_match.getRange(MATCH_INDEX.TIME).setValue(clinic_time);
  sheet_match.getRange(MATCH_INDEX.MANAGERS).setValue(manager_names);

  // Update Match List Sheet file
  var firstName = "";
  var lastName = "";
  var name_row_ndx = 0;
  var actuallyMatched = [];

  Logger.log("Number of rooms");
  Logger.log(num_rooms);

  Logger.log("Number of providers");
  Logger.log(matchList.length);

  var roll_over_providers = [];
  var num_slots = matchList.length < num_rooms ? matchList.length : num_rooms;

  Logger.log("Number of slots");
  Logger.log(num_slots);

  // Fill rooms with people who can see patients alone
  for (i = 0; i < num_slots; i++) {
    // Get index of name on sign up sheet. Repetative but keeps match sorting cleaner
    name_row_ndx = 2;
    for(var j = 0; j < lastRow-1; j++) {
      if (date.valueOf() == sign_dates[j][0].valueOf() && sign_names[j][0] == matchList[i])
        name_row_ndx += j; // List index offset from sheet
    }
    ptsAlone = sheet_sign.getRange(name_row_ndx, SIGN_INDEX.PTS_ALONE).getValue();

    if (ptsAlone == "Yes") {
      actuallyMatched.push(matchList[i]);
      nameArr = findCellByName(matchList[i]);
      firstName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
      lastName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 1).setValue("Room " + (i + 1).toString());
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue(firstName + " " + lastName + ", " + getYearTag(nameArr[0]));
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);
    } else {
      roll_over_providers.push(matchList.splice(i,1)[0]);
      i -= 1;
    }
    if (matchList.length <= (i+1)) {num_slots = matchList.length; break;}
  }

  Logger.log("Roll over providers")
  Logger.log(roll_over_providers)

  // Fill the second room spot
  var matchListP2 = roll_over_providers.concat(matchList.slice(num_slots));
  num_slots2 = matchListP2.length < num_slots ? matchListP2.length : num_slots;
  var prev_name = "";

  Logger.log("Number of slots (for 2nd pass)");
  Logger.log(num_slots2);

  for (i = 0; i < num_slots2; i++) {
    actuallyMatched.push(matchListP2[i]);
    nameArr = findCellByName(matchListP2[i]);
    firstName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
    lastName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();

    prev_name = sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue(prev_name + "\n" + firstName + " " + lastName + ", " + getYearTag(nameArr[0]));
  }

  Logger.log("Match list part 2")
  Logger.log(matchListP2)

  // Add volunteer spaces
  for (i = 0; i < num_slots; i++) {
    prev_name = sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue(prev_name + "\nVolunteer: ");
  }

  // Add DIME Manager slot
  sheet_match.getRange(num_slots + MATCH_INDEX.NAMES, 1).setValue("DIME Managers");
  sheet_match.getRange(num_slots + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange((num_slots+1) + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Add DIME Provider slot
  sheet_match.getRange((num_slots+1) + MATCH_INDEX.NAMES, 1).setValue("DIME Providers");
  sheet_match.getRange((num_slots+1) + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange(num_slots + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Add lay counselor slot
  sheet_match.getRange((num_slots+2) + MATCH_INDEX.NAMES, 1).setValue("Lay Counselors");
  sheet_match.getRange((num_slots+2) + MATCH_INDEX.NAMES, 2).setValue(GET_INFO("LayCouns", "name"));
  sheet_match.getRange((num_slots+2) + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange((num_slots+2) + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // ** Mutable changes after here ** // 

  // Update match stats
  var comments = "";
  var diet_restrict = "";
  var manager_email_body = "";
  for (name in actuallyMatched) {
    nameArr = findCellByName(actuallyMatched[name])
    var tmp = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.MATCHES).getValue();
    if (tmp == "") tmp = 0;
    if (!DEBUG) {
      sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.MATCHES).setValue(parseInt(tmp) + 1);
      sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.DATE).setValue(date);
    } else {
      Logger.log(`DEBUG: Would update TRACKER sheet for ${actuallyMatched[name]}: Matches = ${parseInt(tmp) + 1}, Date = ${date}`);
    }

    // Account for dietary restrictions and comments
    name_row_ndx = 2;
    for(var j = 0; j < lastRow-1; j++) {
      if (date.valueOf() == sign_dates[j][0].valueOf() && sign_names[j][0] == actuallyMatched[name])
        name_row_ndx += j; // List index offset from sheet
    }
    diet_restrict = sheet_sign.getRange(name_row_ndx, SIGN_INDEX.DIET).getValue();
    comments = sheet_sign.getRange(name_row_ndx, SIGN_INDEX.COMMENTS).getValue(); 

    if ((diet_restrict != "None" && diet_restrict != "") || comments != "") {
      manager_email_body += actuallyMatched[name] + " -- Dietary restrictions: " + diet_restrict + "; Comments: " + comments + "\n"; 
    }
  }

  if (manager_email_body == "") {
    manager_email_body = "No comments or dietary restictions noted by matched students.";
  }

  // Send email prompting managers to fill in the number of rooms needed
  var html_body = HtmlService.createTemplateFromFile('PrelimMatchEmail');
  var linkMatch = "https://docs.google.com/spreadsheets/d/" + SHEETS_ID.MATCH + "/edit?usp=sharing";
  html_body.link_match = linkMatch;
  html_body.sign_up_notes = manager_email_body;
  var email_html = html_body.evaluate().getContent();
  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : manager_emails + "," + GET_INFO("DIMEManager", "email") + "," + GET_INFO("LayCouns", "email"),
    subject:  "Notes from SOC sign up",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: email_html,
    name: "SOC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Preliminary match list email sent to Webmaster instead of managers for SOC on ${date}`);
  }

  FormApp.getActiveForm().deleteAllResponses();
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
  var form = FormApp.getActiveForm();
  var sheet = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];
  var sheets_tracker = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  
  // Get the response that was submitted.
  var formResponse = e.response;
  Logger.log(formResponse.getItemResponses()[0].getResponse()); // log name for error checking

  var descr = form.getDescription().split(";");
  var date = descr[0];
  var clinic_type_code = descr[1];
  let lastRow = sheet
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
  
  var itemResponses = formResponse.getItemResponses();
  var name = itemResponses[0].getResponse();
  var nameArr = findCellByName(name);

  // Prevent resubmission 
  var usedNames = sheet.getRange(2, SIGN_INDEX.NAME, lastRow-1).getValues();
  var usedDates = sheet.getRange(2, SIGN_INDEX.DATE, lastRow-1).getValues();
  for(var i = 0; i < lastRow-2; i++) {
    if (name == usedNames[i][0] && 
        new Date(date).valueOf() == usedDates[i][0].valueOf()) {
      Logger.log(name);
      Logger.log("Form resubmission");
      return;
    } else if (isNaN(new Date(date).valueOf())) {
      Logger.log(name);
      Logger.log(date);
      Logger.log("Bad date");
      return;
    }
  }

  // Set the date to the date of the clinic
  sheet.getRange(lastRow, SIGN_INDEX.DATE).setValue(date);
  sheet.getRange(lastRow, SIGN_INDEX.CLINIC_TYPE).setValue(clinic_type_code);

  // Update the sign up counter
  var tmp = sheets_tracker[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.SIGNUPS).getValue();
  if (tmp == "") { 
    tmp = 0;
  }
  if (!DEBUG) {
    sheets_tracker[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.SIGNUPS).setValue(parseInt(tmp) + 1);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Signups = ${parseInt(tmp) + 1}`);
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