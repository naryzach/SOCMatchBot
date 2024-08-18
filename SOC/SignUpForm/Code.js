// Need to update SHEET ID's with new instances (link new sheet to responses first)
// Run createTriggers() after Match Tracker is updated
// Update clinic dates for proper flow of automation

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
  DEFAULT_NUM_ROOMS: "10"
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
      handleLeadTime(form, clinicInfo);
    } else if (clinicDate.valueOf() === checkDates.manage.valueOf()) {
      handleManageTime(spreadsheet, clinicInfo);
    } else if (clinicDate.valueOf() === checkDates.close.valueOf()) {
      handleCloseTime(form, spreadsheet, clinicInfo);
    }
  }
}

/**
 * Handles the lead time for clinic sign-ups.
 * 
 * This function is called when it's time to open sign-ups for a clinic.
 * It performs the following tasks:
 * 1. Updates the list of students who can sign up
 * 2. Sets up the sign-up form with clinic details
 * 3. Calculates the closing date for sign-ups
 * 4. Sends a sign-up email to students
 * 
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form object for sign-ups
 * @param {Object} clinicInfo - Object containing clinic details (date, type, time, etc.)
 */
function handleLeadTime(form, clinicInfo) {
  updateStudents();
  form.setTitle(`${clinicInfo.type} Clinic Sign Up -- ${clinicInfo.dateString} from ${clinicInfo.time}`);
  form.setDescription(`${Utilities.formatDate(clinicInfo.date, 'GMT', 'MM/dd/YYYY')};${clinicInfo.typeCode}`);
  form.setAcceptingResponses(true);

  const closeDate = new Date();
  closeDate.setDate(closeDate.getDate() + (SIGNUP_DAYS.LEAD - SIGNUP_DAYS.CLOSE));

  // Send sign-up email
  const htmlBody = HtmlService.createTemplateFromFile('SignUpEmail');
  htmlBody.type = clinicInfo.type;
  htmlBody.date = clinicInfo.dateString;
  htmlBody.close_date = Utilities.formatDate(closeDate, 'GMT', 'EEEE, MMMM dd, YYYY');
  htmlBody.time = clinicInfo.time;
  htmlBody.link = `https://docs.google.com/forms/d/e/${FORMS_ID.OFFICIAL}/viewform?usp=sf_link`;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");

  const emailHtml = htmlBody.evaluate().getContent();
  const recipient = DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("ClassLists", "email");

  MailApp.sendEmail({
    to: recipient,
    subject: `Sign up for SOC on ${clinicInfo.dateString}`,
    replyTo: clinicInfo.managerEmail,
    htmlBody: emailHtml,
    name: "SOC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Sign-up email sent to Webmaster instead of class lists for SOC on ${clinicInfo.dateString}`);
  }
}

/**
 * Handles the management time for clinic preparation.
 * 
 * This function is called a few days before the clinic to prepare for room allocation.
 * It performs the following tasks:
 * 1. Sets the default number of rooms for the clinic
 * 2. Sends an email to the clinic manager about room availability
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The Google Spreadsheet object
 * @param {Object} clinicInfo - Object containing clinic details (date, type, time, etc.)
 */
function handleManageTime(spreadsheet, clinicInfo) {
  spreadsheet.getRange(DATE_INDEX.ROOMS).setValue(DATE_INDEX.DEFAULT_NUM_ROOMS);

  // Send room availability email
  const htmlBody = HtmlService.createTemplateFromFile('RoomNumEmail');
  htmlBody.type = clinicInfo.type;
  htmlBody.date = clinicInfo.dateString;
  htmlBody.cell_ndx = DATE_INDEX.ROOMS;
  htmlBody.time = clinicInfo.time;
  htmlBody.link_date = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.DATES}/edit?usp=sharing`;
  htmlBody.link_match = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`;
  htmlBody.link_match_track = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`;
  htmlBody.link_form_mod = `https://docs.google.com/forms/d/e/${FORMS_ID.MOD}/viewform?usp=sf_link`;

  const emailHtml = htmlBody.evaluate().getContent();
  const recipient = DEBUG ? GET_INFO("Webmaster", "email") : clinicInfo.managerEmail;

  MailApp.sendEmail({
    to: recipient,
    subject: "SOC Clinic Room Availability Form",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: emailHtml,
    name: "SOC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Room availability email sent to Webmaster instead of managers for SOC on ${clinicInfo.dateString}`);
  }
}

/**
 * Handles the closing time for clinic sign-ups.
 * 
 * This function is called when it's time to close sign-ups for a clinic.
 * It performs the following tasks:
 * 1. Closes the sign-up form
 * 2. Updates the match list based on sign-ups and available rooms
 * 3. Stores temporary data for later use
 * 4. Schedules a trigger to send the match list email later
 * 
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form object for sign-ups
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The Google Spreadsheet object
 * @param {Object} clinicInfo - Object containing clinic details (date, type, time, etc.)
 */
function handleCloseTime(form, spreadsheet, clinicInfo) {
  form.setTitle("Sign Ups Closed.");
  form.setDescription("Thank you for your interest. Please check back when another clinic is closer.");
  form.setAcceptingResponses(false);

  const numRooms = parseInt(spreadsheet.getRange(DATE_INDEX.ROOMS).getValue());
  updateMatchList(clinicInfo.date, clinicInfo.typeCode, numRooms);

  // Store temporary data
  const tempDataRange = spreadsheet.getRange("X1:X4");
  tempDataRange.setValues([
    [clinicInfo.type],
    [clinicInfo.dateString],
    [clinicInfo.time],
    [clinicInfo.managerEmail]
  ]);

  // Schedule match list delay
  const today = new Date();
  const triggerTime = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 17, 0);
  
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'matchListDelay') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('matchListDelay')
    .timeBased()
    .at(triggerTime)
    .create();
}

/**
 * Handles the delayed sending of the match list.
 * This function is triggered at a specific time after sign-ups close.
 * 
 * It performs the following tasks:
 * 1. Retrieves clinic information from temporary storage
 * 2. Generates a PDF of the match list
 * 3. Sends an email with the match list to all relevant parties
 */
function matchListDelay() {
  // Grab tmp data (from date sheet)
  var spreadsheet = SpreadsheetApp.openById(SHEETS_ID.DATES);
  var type = spreadsheet.getRange("X1").getValue().toString();
  var raw_date = new Date(spreadsheet.getRange("X2").getValue().toString());
  var tz = "GMT-" + String(raw_date.getTimezoneOffset() / 60);
  var date_string = Utilities.formatDate(raw_date, tz, 'EEEE, MMMM dd, YYYY');
  var time = spreadsheet.getRange("X3").getValue().toString();
  var clinic_emails = spreadsheet.getRange("X4").getValue().toString();

  var file = makeMatchPDF(raw_date); // make the PDF of the match list

  // Format email from HTML file
  var html_body = HtmlService.createTemplateFromFile('MatchEmail');  
  html_body.type = type;
  html_body.date = date_string;
  html_body.time = time;
  html_body.feedback_email = GET_INFO("Webmaster", "email");
  var email_html = html_body.evaluate().getContent();
  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("ClassLists", "email"),
    subject:  "Match list for SOC on " + date_string,
    replyTo: clinic_emails,
    htmlBody: email_html,
    attachments: [file.getAs(MimeType.PDF)],
    name: "SOC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Final match list email sent to Webmaster instead of class lists for SOC on ${date_string}`);
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
 * Generates a PDF of the match list for a given date.
 * 
 * @param {Date} date - The date of the clinic
 * @returns {GoogleAppsScript.Drive.File} The generated PDF file
 */
function makeMatchPDF(date) {
  // PDF Creation https://developers.google.com/apps-script/samples/automations/generate-pdfs
  pdfName = "MatchList" + (date).toISOString().split('T')[0] + ".pdf"
  var sheet = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];

  const fr = 0, fc = 0, lc = 4, lr = 30;
  const url = "https://docs.google.com/spreadsheets/d/" + SHEETS_ID.MATCH + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.25&" +
    "bottom_margin=0.25&" +
    "left_margin=0.25&" +
    "right_margin=0.25&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName);

  // Gets the folder in Drive where the PDFs are stored.
  const folder = DriveApp.getFoldersByName("MatchListsSOC").next();

  // Not entirely sure of this is necessary or if the next file query is
  const pdfFile = folder.createFile(blob);
  //return pdfFile;

  var file = DriveApp.getFilesByName(pdfName).next();

  return file;
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
    Logger.log("DEBUG: Would update TRACKER sheet for " + name + ": Signups = " + (parseInt(tmp) + 1));
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

/**
 * Builds and returns a sorted list of student names from the tracker spreadsheet.
 * 
 * This function performs the following tasks:
 * 1. Retrieves all sheets from the tracker spreadsheet
 * 2. Iterates through each sheet, collecting student names
 * 3. Formats each name as "Last, First (YearTag)"
 * 4. Checks for and logs any duplicate names
 * 5. Returns a sorted list of unique student names
 *
 * @returns {string[]} A sorted array of formatted student names
 */
function buildNameList() {
  const sheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const studentNames = new Set();

  sheets.forEach((sheet, sheetIndex) => {
    const yearTag = getYearTag(sheetIndex);
    if (!yearTag) return;

    const lastRow = sheet.getLastRow();
    const lastNames = sheet.getRange(2, TRACK_INDEX.LASTNAME, lastRow - 1, 1).getValues();
    const firstNames = sheet.getRange(2, TRACK_INDEX.FIRSTNAME, lastRow - 1, 1).getValues();

    lastNames.forEach((lastName, index) => {
      if (lastName[0] !== "") {
        const newName = `${lastName[0]}, ${firstNames[index][0]} (${yearTag})`;
        if (studentNames.has(newName)) {
          Logger.log(`Duplicate: ${newName}`);
        } else {
          studentNames.add(newName);
        }
      }
    });
  });

  const sortedNames = Array.from(studentNames).sort();
  Logger.log(sortedNames);
  return sortedNames;
}

/**
 * Finds the sheet index and row index for a given student name.
 * 
 * @param {string} name - The formatted name of the student (e.g., "Last, First (MS1)").
 * @returns {number[]} An array containing [sheetIndex, rowIndex] of the student's entry.
 */
function findCellByName(name) {
  const sheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const [lastName, firstName] = name.slice(0, -6).split(", ");

  for (let sheetIndex = 0; sheetIndex < sheets.length; sheetIndex++) {
    const sheet = sheets[sheetIndex];
    const lastNames = sheet.getRange(2, TRACK_INDEX.LASTNAME, sheet.getLastRow() - 1, 1).getValues();
    const firstNames = sheet.getRange(2, TRACK_INDEX.FIRSTNAME, sheet.getLastRow() - 1, 1).getValues();

    const rowIndex = lastNames.findIndex((row, index) => 
      row[0] === lastName && firstNames[index][0] === firstName
    );

    if (rowIndex !== -1) {
      return [sheetIndex, rowIndex + 2]; // +2 because we start from row 2 and array is 0-indexed
    }
  }

  Logger.log(`Did not find name: ${name}`);
  return [-1, -1];
}

/**
 * Updates the list of student names in the Google Form.
 * 
 * This function retrieves the current list of student names and
 * updates the corresponding form item with these names as choices.
 */
function updateStudents() {
  const form = FormApp.getActiveForm();
  const namesList = form.getItemById(NAMES_ID).asListItem();
  const studentNames = buildNameList();
  namesList.setChoiceValues(studentNames);
}

/**
 * Returns the year tag based on the sheet index.
 * 
 * @param {number} sheetNum - The index of the sheet.
 * @returns {string|number} The year tag (e.g., "MS1", "PA2") or 0 if invalid.
 */
function getYearTag(sheetNum) {
  const tags = ["MS1", "MS2", "MS3", "MS4", "PA1", "PA2"];
  return sheetNum < tags.length ? tags[sheetNum] : 0;
}

/**
 * Retrieves information about a specific position from the People sheet.
 * 
 * @param {string} position - The position to look up (e.g., "CEO", "ROCManager", "ClassLists").
 * @param {string} info - The type of information to retrieve ("name" or "email").
 * @returns {string} The requested information (name or email) for the specified position.
 * 
 * This function:
 * 1. Opens the People sheet using the SHEET_PEOPLE ID.
 * 2. Uses a switch statement to find the correct row for the given position.
 * 3. Retrieves the name and email from the appropriate cells.
 * 4. Returns the requested information (name or email) based on the 'info' parameter.
 * 
 * If the position is not found or the info type is invalid, it returns an error message.
 */
function GET_INFO(position, info) {
  const sheet = SpreadsheetApp.openById(SHEETS_ID.PEOPLE).getSheets()[0];
  const positions = {
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

  const rowIndex = positions[position];
  if (!rowIndex) {
    return info === "email" ? "Email Not Found" : "Name Not Found";
  }

  const name = sheet.getRange(rowIndex, 2).getValue();
  const email = sheet.getRange(rowIndex, 3).getValue();

  return info.toLowerCase() === "email" ? email : name;
}