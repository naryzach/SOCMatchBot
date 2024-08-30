/**
 * Student Outreach Clinic (SOC) Sign-Up Form Handlers
 * 
 * This script contains handler functions for managing the sign-up process,
 * preliminary and final matching, and email communications for the Student Outreach Clinic program.
 * It works in conjunction with Google Sheets and Forms to automate the scheduling process.
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
 * - Ensure this file is used in conjunction with the main SOC scheduling script
 * - Update email recipients and content as needed for your specific SOC program
 * - Verify that all referenced spreadsheet IDs and ranges are correct
 */

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
function handleSignUp(form, clinicInfo) {
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
function handleManagers(spreadsheet, clinicInfo) {
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
function handleFinalMatch(form, spreadsheet, clinicInfo) {
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

  if (!DEBUG) {
    ScriptApp.newTrigger('matchListDelay')
      .timeBased()
      .at(triggerTime)
      .create();
  }
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
* Generates a match list based on student sign-ups and their scores.
* 
* @param {Date} date - The date of the clinic.
* @param {number} num_rooms - The number of available rooms for the clinic.
* @returns {string[]} An array of student names representing the match list.
* 
* This function performs the following tasks:
* 1. Retrieves necessary data from tracker and sign-up sheets.
* 2. Calculates a match score for each student based on various factors:
*    - Number of sign-ups and matches
*    - SOC membership status
*    - Fourth-year elective status
*    - Seniority
*    - Time since last match
*    - Cancellation history
* 3. Sorts students based on their match scores.
* 4. Returns a list of matched students, limited by the number of available rooms.
*/
function generateMatchList(date, num_rooms) {
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  // Get the last row of the sign-up sheet
  const lastRow = sheetSign.getLastRow();
  // Get the sign-up dates and names for all students
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow).getValues();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow).getValues();

  // Gather names of signups for current dated clinic
  const largeNameList = signDates.slice(0, lastRow - 1)
    .map((dateValue, index) => dateValue[0].valueOf() === date.valueOf() ? signNames[index][0] : null)
    .filter(name => name !== null);

  const namesWithScores = {};

  // Generate match list
  for (const name of largeNameList) {
    const nameRowNdx = signNames.findIndex(row => row[0] === name && signDates[signNames.indexOf(row)][0].valueOf() === date.valueOf()) + 2;
    const nameArr = findCellByName(name);
    
    // Check for cancelled names
    if (nameArr[0] === -1) {
      Logger.log(`Name error: ${name}`);
      if (name.endsWith("CXL")) {
        const originalName = name.slice(0, -3);
        const cancelledNameArr = findCellByName(originalName);
        if (cancelledNameArr[0] !== -1) {
          let tmp = sheetsTrack[cancelledNameArr[0]].getRange(cancelledNameArr[1] + 1, TRACK_INDEX.CXLEARLY).getValue();
          tmp = tmp === "" ? 0 : parseInt(tmp);
          if (!DEBUG) {
            sheetsTrack[cancelledNameArr[0]].getRange(cancelledNameArr[1] + 1, TRACK_INDEX.CXLEARLY).setValue(tmp + 1);
          } else {
            Logger.log(`DEBUG: Would update TRACKER sheet for ${originalName}: CXLEARLY = ${tmp + 1}`);
          }
        }
      }
      continue;
    }
    
    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1];
    const studentData = {
      signUps: parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.SIGNUPS).getValue()) || 0,
      matches: parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.MATCHES).getValue()) || 0,
      noShow: parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.NOSHOW).getValue()) || 0,
      cxlLate: parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.CXLLATE).getValue()) || 0,
      cxlEarly: parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.CXLEARLY).getValue()) || 0,
      lastDate: trackSheet.getRange(trackRow, TRACK_INDEX.DATE).getValue(),
      fourthYrElect: sheetSign.getRange(nameRowNdx, SIGN_INDEX.ELECTIVE).getValue(),
      socPos: sheetSign.getRange(nameRowNdx, SIGN_INDEX.SOC_POS).getValue()
    };

    // Calculate match score
    let matchScore = studentData.signUps - studentData.matches;

    // Adjust score based on student status and position
    if (studentData.socPos == "Yes" && nameArr[0] <= 1) { // SOC members (MS1/2s)
      matchScore *= 2;
    }
    if (studentData.fourthYrElect == "Yes" && nameArr[0] == 3) { // MS4s on elective
      matchScore += 500;
    }

    // Add points based on seniority
    const seniorityPoints = [0, 50, 500, 1000, 0, 0];
    matchScore += seniorityPoints[nameArr[0]] || 0;

    // Adjust for last match date
    if (studentData.lastDate == "") {
      matchScore += 25; // Never been matched
    } else {
      const daysSince = (new Date() - new Date(studentData.lastDate)) / (1000 * 60 * 60 * 24);
      matchScore += daysSince / 365;
    }

    // Apply cancellation penalties
    matchScore -= studentData.noShow * 3 + studentData.cxlLate * 2 + studentData.cxlEarly;

    // Create dictionary of name (key) and score (value)
    namesWithScores[name] = matchScore;
  }

  Logger.log("Names with scores:")
  Logger.log(namesWithScores);

  // Generate match list based on points
  const sortedNames = Object.entries(namesWithScores)
    .sort((a, b) => b[1] - a[1])
    .map(entry => entry[0]);

  const matchList = sortedNames.slice(0, Math.min(sortedNames.length, num_rooms * 2));

  Logger.log("Prelim match list:");
  Logger.log(matchList);

  return matchList;
}

/**
* Sets up the match list in the Google Sheet.
* 
* @param {string[]} matchList - Array of student names who have been matched.
* @param {string} clinicTime - The time of the clinic.
* @param {Object} clinicInfo - Object containing clinic information (title, managerType).
* @param {Date} date - The date of the clinic.
* @param {number} num_rooms - The number of available rooms for the clinic.
* @returns {string[]} An array of student names who were actually matched.
* 
* This function performs the following tasks:
* 1. Clears and resets the match list sheet.
* 2. Populates the sheet with clinic information (title, date, time, managers).
* 3. Fills in matched students, prioritizing those who can see patients alone.
* 4. Adds slots for volunteers, DIME managers, DIME providers, and lay counselors.
* 5. Applies formatting to the sheet (borders, text wrapping).
*/
function setupMatchList(matchList, clinicTime, clinicInfo, date, num_rooms) {
  const sheetMatch = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  // Get the last row of the sign-up sheet
  const lastRow = sheetSign.getLastRow();
  // Get the sign-up dates and names for all students
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow).getValues();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow).getValues();

  // Clear Match List Sheet file names and remove borders
  sheetMatch.getRange(MATCH_INDEX.NAMES, 1, 25, 3).clearContent().setBorder(false, false, false, false, false, false);

  // Clear physicians and chalk talk
  sheetMatch.getRange(MATCH_INDEX.PHYS1).clearContent();
  sheetMatch.getRange(MATCH_INDEX.PHYS2).clearContent();
  sheetMatch.getRange(MATCH_INDEX.CHALK_TALK).clearContent();

  // Get clinic title and manager names
  const clinicTitle = clinicInfo.title;
  const managerNames = GET_INFO(clinicInfo.managerType, "name");

  sheetMatch.getRange(MATCH_INDEX.TITLE).setValue(clinicTitle);
  sheetMatch.getRange(MATCH_INDEX.DATE).setValue(date);
  sheetMatch.getRange(MATCH_INDEX.TIME).setValue(clinicTime);
  sheetMatch.getRange(MATCH_INDEX.MANAGERS).setValue(managerNames);
  sheetMatch.getRange(MATCH_INDEX.LIAISON).setValue(GET_INFO("Liaison", "name"));
  
  // Update Match List Sheet file
  // Initialize variables
  let firstName, lastName;
  const actuallyMatched = [];
  const rollOverProviders = [];

  Logger.log(`Number of rooms: ${num_rooms}`);
  Logger.log(`Number of providers: ${matchList.length}`);

  let numSlots = Math.min(matchList.length, num_rooms);
  Logger.log(`Number of slots: ${numSlots}`);

  // Fill rooms with people who can see patients alone
  for (let i = 0; i < numSlots; i++) {
    const nameRowNdx = signNames.findIndex(n => n[0] === matchList[i] && signDates[signNames.indexOf(n)][0].valueOf() === date.valueOf()) + 2;
    
    const ptsAlone = sheetSign.getRange(nameRowNdx, SIGN_INDEX.PTS_ALONE).getValue();

    if (ptsAlone === "Yes") {
      actuallyMatched.push(matchList[i]);
      const nameArr = findCellByName(matchList[i]);
      const trackSheet = sheetsTrack[nameArr[0]];
      const trackRow = nameArr[1];
      firstName = trackSheet.getRange(trackRow, TRACK_INDEX.FIRSTNAME).getValue();
      lastName = trackSheet.getRange(trackRow, TRACK_INDEX.LASTNAME).getValue();
      
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1).setValue(`Room ${i + 1}`);
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);
    } else {
      rollOverProviders.push(matchList.splice(i, 1)[0]);
      i--;
    }
    if (matchList.length <= (i+1)) {numSlots = matchList.length; break;}
  }

  Logger.log("Roll over providers:");
  Logger.log(rollOverProviders);

  // Fill the second room spot
  const matchListP2 = rollOverProviders.concat(matchList.slice(numSlots));
  const numSlots2 = Math.min(matchListP2.length, numSlots);

  Logger.log(`Number of slots (for 2nd pass): ${numSlots2}`);

  for (let i = 0; i < numSlots2; i++) {
    actuallyMatched.push(matchListP2[i]);
    const nameArr = findCellByName(matchListP2[i]);
    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1];
    firstName = trackSheet.getRange(trackRow, TRACK_INDEX.FIRSTNAME).getValue();
    lastName = trackSheet.getRange(trackRow, TRACK_INDEX.LASTNAME).getValue();

    const prevName = sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
    sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${prevName}\n${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
  }

  Logger.log("Match list part 2:");
  Logger.log(matchListP2);

  // Add volunteer spaces
  for (let i = 0; i < numSlots; i++) {
    const prevName = sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
    sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${prevName}\nVolunteer: `);
  }

  // Add DIME Manager slot
  sheetMatch.getRange(numSlots + MATCH_INDEX.NAMES, 1).setValue("DIME Managers");
  sheetMatch.getRange(numSlots + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheetMatch.getRange(numSlots + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Add DIME Provider slot
  sheetMatch.getRange(numSlots + 1 + MATCH_INDEX.NAMES, 1).setValue("DIME Providers");
  sheetMatch.getRange(numSlots + 1 + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheetMatch.getRange(numSlots + 1 + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Add lay counselor slot
  sheetMatch.getRange(numSlots + 2 + MATCH_INDEX.NAMES, 1).setValue("Lay Counselors");
  sheetMatch.getRange(numSlots + 2 + MATCH_INDEX.NAMES, 2).setValue(GET_INFO("LayCouns", "name"));
  sheetMatch.getRange(numSlots + 2 + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheetMatch.getRange(numSlots + 2 + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  return actuallyMatched;
}

/**
* Updates match statistics and sends an email with match information to clinic managers.
* 
* @param {string[]} actuallyMatched - Array of names of students who were matched for the clinic.
* @param {Date} date - The date of the clinic.
* 
* This function performs the following tasks:
* 1. Updates the tracker sheets with new match information for each matched student:
*    - Increments the match count
*    - Updates the last match date
*    - Appends the new date to the list of all match dates
* 2. Gathers dietary restrictions and comments from the sign-up sheet for matched students
* 3. Prepares and sends an email to clinic managers (or Webmaster in DEBUG mode) containing:
*    - A link to the match list spreadsheet
*    - Notes about dietary restrictions and comments from matched students
* 4. Deletes all responses from the active form after processing
* 
* The function handles both DEBUG and normal operation modes, logging actions instead of 
* making changes in DEBUG mode.
*/
function updateMatchStats(actuallyMatched, clinicInfo, date) {
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  // Get the last row of the sign-up sheet
  const lastRow = sheetSign.getLastRow();
  // Get the sign-up dates and names for all students
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow).getValues();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow).getValues();
  
  const managerEmails = GET_INFO(clinicInfo.managerType, "email");
  let managerEmailBody = "";
  for (const name of actuallyMatched) {
    const nameArr = findCellByName(name);
    const trackSheet = sheetsTrack[nameArr[0]];
    const row = nameArr[1];

    // Update match count and date
    let matches = trackSheet.getRange(row, TRACK_INDEX.MATCHES).getValue() || 0;
    if (!DEBUG) {
      trackSheet.getRange(row, TRACK_INDEX.MATCHES).setValue(parseInt(matches) + 1);
      trackSheet.getRange(row, TRACK_INDEX.DATE).setValue(date);
      const allDates = trackSheet.getRange(row, TRACK_INDEX.DATE_ALL).getValue();
      trackSheet.getRange(row, TRACK_INDEX.DATE_ALL).setValue(allDates ? allDates + "," + date : date);
    } else {
      Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Matches = ${parseInt(matches) + 1}, Date = ${date}`);
    }

    // Gather sign-up information
    const nameRowIndex = signNames.findIndex(n => n[0] === name && signDates[signNames.indexOf(n)][0].valueOf() === date.valueOf()) + 2;
    const dietRestrict = sheetSign.getRange(nameRowIndex, SIGN_INDEX.DIET).getValue();
    const comments = sheetSign.getRange(nameRowIndex, SIGN_INDEX.COMMENTS).getValue();

    if (dietRestrict !== "None" && dietRestrict !== "" || comments !== "") {
      managerEmailBody += `${name} -- Dietary restrictions: ${dietRestrict}; Comments: ${comments}\n`;
    }
  }

  if (managerEmailBody === "") {
    managerEmailBody = "No comments or dietary restrictions noted by matched students.";
  }

  // Send email prompting managers to fill in the number of rooms needed
  const htmlBody = HtmlService.createTemplateFromFile('PrelimMatchEmail');
  const linkMatch = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`;
  htmlBody.link_match = linkMatch;
  htmlBody.sign_up_notes = managerEmailBody;
  const emailHtml = htmlBody.evaluate().getContent();
  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : `${managerEmails},${GET_INFO("DIMEManager", "email")},${GET_INFO("LayCouns", "email")},${GET_INFO("Liaison", "email")}`,
    subject: "Notes from SOC sign up",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: emailHtml,
    name: "SOC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Preliminary match list email sent to Webmaster instead of managers for SOC on ${date}`);
  } else {
    FormApp.getActiveForm().deleteAllResponses();
  }
}

/**
* Generates a PDF of the match list for a given date.
* 
* @param {Date} date - The date of the clinic
* @returns {GoogleAppsScript.Drive.File} The generated PDF file
*/
function makeMatchPDF(date) {
  // PDF Creation https://developers.google.com/apps-script/samples/automations/generate-pdfs
  const pdfName = "MatchList" + (date).toISOString().split('T')[0] + ".pdf";
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