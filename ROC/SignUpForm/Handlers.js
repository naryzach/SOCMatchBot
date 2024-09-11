/**
 * Rural Outreach Clinic (ROC) Sign-Up Form Handlers
 * 
 * This script contains handler functions for managing the sign-up process,
 * preliminary and final matching, and email communications for the Rural Outreach Clinic program.
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
 * - Ensure this file is used in conjunction with the main ROC scheduling script
 * - Update email recipients and content as needed for your specific ROC program
 * - Verify that all referenced spreadsheet IDs and ranges are correct
 */

/**
 * Handles the sign-up process for a clinic.
 * Updates the form, sends sign-up emails, and manages form responses.
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form to update.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {string} timeZone - Time zone string for date formatting.
 * @param {Object} clinicInfo - Object containing clinic information.
 * @param {Object} links - Object containing various relevant URLs.
 */
function handleSignUp(form, dateString, clinicDate, timeZone, clinicInfo, links) {
  updateStudents();
  form.setTitle(`${clinicInfo.type} Clinic Sign Up -- ${dateString} from ${clinicInfo.time}`);
  form.setDescription(`${Utilities.formatDate(clinicDate, timeZone, 'MM/dd/YYYY')};${clinicInfo.typeCode}`);
  form.setAcceptingResponses(true);

  const formCloseDate = new Date();
  formCloseDate.setDate(formCloseDate.getDate() + (SIGNUP_DAYS.LEAD - SIGNUP_DAYS.MANAGE));

  const htmlBody = HtmlService.createTemplateFromFile('SignUpEmail');
  htmlBody.type = clinicInfo.type;
  htmlBody.date = dateString;
  htmlBody.close_date = Utilities.formatDate(formCloseDate, timeZone, 'EEEE, MMMM dd, YYYY');
  htmlBody.time = clinicInfo.time;
  htmlBody.link = links.form;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");
  const emailHtml = htmlBody.evaluate().getContent();

  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("ClassLists", "email"),
    subject: `Sign up for ROC on ${dateString}`,
    replyTo: GET_INFO("ROCManager", "email"),
    htmlBody: emailHtml,
    name: "ROC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Sign-up email sent to Webmaster instead of class lists for ROC on ${dateString}`);
  }
}

/**
 * Handles the preliminary match process for a clinic.
 * Closes the sign-up form and updates the match list.
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form to update.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {Object} clinicInfo - Object containing clinic information.
 */
function handlePreliminaryMatch(form, dateString, clinicDate, clinicInfo) {
  form.setTitle("Sign Ups Closed.");
  form.setDescription("Thank you for your interest. Please check back when another clinic is closer.");
  form.setAcceptingResponses(false);
  updateMatchList(clinicDate, clinicInfo.typeCode, clinicInfo.rooms, clinicInfo.address);
}

/**
 * Handles the final match process for a clinic.
 * Creates and sends the final match PDF to participants.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {Object} clinicInfo - Object containing clinic information.
 */
function handleFinalMatch(dateString, clinicDate, clinicInfo) {
  const file = makeMatchPDF(clinicDate, clinicInfo.typeCode);
  
  const htmlBody = HtmlService.createTemplateFromFile('MatchEmail');
  htmlBody.type = clinicInfo.type;
  htmlBody.date = dateString;
  htmlBody.time = clinicInfo.time;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");
  const emailHtml = htmlBody.evaluate().getContent();

  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("ClassLists", "email"),
    subject: `Match list for ROC on ${dateString}`,
    replyTo: GET_INFO("ROCManager", "email"),
    htmlBody: emailHtml,
    attachments: [file.getAs(MimeType.PDF)],
    name: "ROC Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Final match list email sent to Webmaster instead of class lists for ROC on ${dateString}`);
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
        let tmp = sheetsTrack[cancelledNameArr[0]].getRange(cancelledNameArr[1], TRACK_INDEX.CXLEARLY).getValue();
        tmp = tmp === "" ? 0 : parseInt(tmp);
        if (!DEBUG) {
          sheetsTrack[cancelledNameArr[0]].getRange(cancelledNameArr[1], TRACK_INDEX.CXLEARLY).setValue(tmp + 1);
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
    socPos: sheetSign.getRange(nameRowNdx, SIGN_INDEX.SOC_POS).getValue(),
    spanish: sheetSign.getRange(nameRowNdx, SIGN_INDEX.SPANISH).getValue()
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

  // Add points for Spanish
  if (studentData.spanish === "Yes") {
    matchScore += 35;
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
function setupMatchList(matchList, clinicInfo, date, num_rooms, address) {
const sheetMatch = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];
const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

// Get the last row of the sign-up sheet
const lastRow = sheetSign.getLastRow();
// Get the sign-up dates and names for all students
const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow).getValues();
const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow).getValues();

// Clear Match List Sheet
const clearRange = sheetMatch.getRange(MATCH_INDEX.NAMES, 1, 25, 3);
clearRange.clearContent().setBorder(false, false, false, false, false, false);

// Clear specific fields
const fieldsToClean = [MATCH_INDEX.CHALK, MATCH_INDEX.INTERPRET, MATCH_INDEX.SHADOW, MATCH_INDEX.PHYS, MATCH_INDEX.VOLUNT];
fieldsToClean.forEach(field => sheetMatch.getRange(field).clearContent());

// Update Match List Sheet header
sheetMatch.getRange(MATCH_INDEX.TITLE).setValue(clinicInfo.title);
sheetMatch.getRange(MATCH_INDEX.DATE).setValue(date);
sheetMatch.getRange(MATCH_INDEX.TIME).setValue(clinicInfo.time);
sheetMatch.getRange(MATCH_INDEX.ADDRESS).setValue(address);

// Update Match List Sheet
let firstName, lastName, nameRowIndex;
const actuallyMatched = [];
const rollOverProviders = [];

Logger.log(`Number of rooms: ${num_rooms}`);
Logger.log(`Number of providers: ${matchList.length}`);

num_rooms -= 1; // DIME takes a room space
let numSlots = Math.min(matchList.length, num_rooms);

Logger.log(`Number of slots: ${numSlots}`);

// Fill rooms with people who can see patients alone
for (let i = 0; i < numSlots; i++) {
  // Find the index of the name on the sign-up sheet
  nameRowIndex = 2;
  for (let j = 0; j < lastRow - 1; j++) {
    if (date.valueOf() == signDates[j][0].valueOf() && signNames[j][0] == matchList[i]) {
      nameRowIndex += j; // List index offset from sheet
      break;
    }
  }
  const ptsAlone = sheetSign.getRange(nameRowIndex, SIGN_INDEX.PTS_ALONE).getValue();

  if (ptsAlone === "Yes") {
    actuallyMatched.push(matchList[i]);
    const nameArr = findCellByName(matchList[i]);
    firstName = sheetsTrack[nameArr[0]].getRange(nameArr[1], TRACK_INDEX.FIRSTNAME).getValue();
    lastName = sheetsTrack[nameArr[0]].getRange(nameArr[1], TRACK_INDEX.LASTNAME).getValue();
    
    // Update match list sheet with provider information
    sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1).setValue(`Room ${i + 1}`);
    sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
    sheetMatch.getRange(i + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
    sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);
  } else {
    rollOverProviders.push(matchList.splice(i, 1)[0]);
    i--; // Adjust index since we removed an item
  }
  if (matchList.length <= (i+1)) {numSlots = matchList.length; break;}
}

Logger.log("Roll over providers:");
Logger.log(rollOverProviders);

// Fill the second room spot
const matchListP2 = rollOverProviders.concat(matchList.slice(numSlots));
const numSlots2 = Math.min(matchListP2.length, numSlots);

Logger.log(`Number of slots (for 2nd pass): ${numSlots2}`);

// Add second provider to each room
for (let i = 0; i < numSlots2; i++) {
  actuallyMatched.push(matchListP2[i]);
  const nameArr = findCellByName(matchListP2[i]);
  firstName = sheetsTrack[nameArr[0]].getRange(nameArr[1], TRACK_INDEX.FIRSTNAME).getValue();
  lastName = sheetsTrack[nameArr[0]].getRange(nameArr[1], TRACK_INDEX.LASTNAME).getValue();

  const prevName = sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
  sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${prevName}\n${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
}

Logger.log(`Match list part 2: ${matchListP2}`);

// Add post-bac spaces
for (let i = 0; i < numSlots + 1; i++) { // Add room back for DIME
  const prevName = sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
  sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${prevName}\nPost-bac: `);
}

// Add DIME slot
sheetMatch.getRange(numSlots + MATCH_INDEX.NAMES, 1).setValue("DIME Providers");
sheetMatch.getRange(numSlots + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
sheetMatch.getRange(numSlots + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

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

let managerEmailBody = "";
const signUpInfo = {};

for (const name of actuallyMatched) {
  const nameArr = findCellByName(name);
  const trackSheet = sheetsTrack[nameArr[0]];
  const row = nameArr[1];

  // Update match count and date
  if (!DEBUG) {
    let matches = trackSheet.getRange(row, TRACK_INDEX.MATCHES).getValue() || 0;
    trackSheet.getRange(row, TRACK_INDEX.MATCHES).setValue(matches + 1);
    trackSheet.getRange(row, TRACK_INDEX.DATE).setValue(date);
    const allDates = trackSheet.getRange(row, TRACK_INDEX.DATE_ALL).getValue();
    trackSheet.getRange(row, TRACK_INDEX.DATE_ALL).setValue(allDates ? allDates + "," + date : date);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Matches incremented, Date set to ${date}`);
  }

  // Gather sign-up information
  const nameRowIndex = signNames.findIndex(n => n[0] === name && signDates[signNames.indexOf(n)][0].valueOf() === date.valueOf()) + 2;
  if (nameRowIndex > 1) {
    signUpInfo[name] = {
      spanish: sheetSign.getRange(nameRowIndex, SIGN_INDEX.SPANISH).getValue(),
      follow: sheetSign.getRange(nameRowIndex, SIGN_INDEX.FOLLOW).getValue(),
      carpool: sheetSign.getRange(nameRowIndex, SIGN_INDEX.CARPOOL).getValue(),
      comments: sheetSign.getRange(nameRowIndex, SIGN_INDEX.COMMENTS).getValue()
    };

    managerEmailBody += `${name} -- Speaks Spanish: ${signUpInfo[name].spanish}; Can have followers: ${signUpInfo[name].follow}; Carpool status: ${signUpInfo[name].carpool}; Comments: ${signUpInfo[name].comments}\n`;
  }
}

// Send email with the preliminary match list for Managers to update
const htmlBody = HtmlService.createTemplateFromFile('PrelimMatchEmail');
const timeZone = `GMT-${date.getTimezoneOffset() / 60}`; // Note: This won't work east of Prime Meridian
const dateString = Utilities.formatDate(date, timeZone, 'EEEE, MMMM dd, YYYY');
const linkMatch = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`;
const linkTrack = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`;

htmlBody.type = clinicInfo.title;
htmlBody.date = dateString;
htmlBody.time = clinicInfo.time;
htmlBody.link = linkMatch;
htmlBody.link_track = linkTrack;
htmlBody.sign_up_notes = managerEmailBody;

MailApp.sendEmail({
  to: DEBUG ? GET_INFO("Webmaster", "email") : `${GET_INFO("ROCManager", "email")},${GET_INFO("DIMEManager", "email")},${GET_INFO("LayCouns", "email")}`,
  subject: "ROC Match List (Prelim) and Notes from Sign-ups",
  replyTo: GET_INFO("Webmaster", "email"),
  htmlBody: htmlBody.evaluate().getContent(),
  name: "ROC Scheduling Assistant"
});

if (DEBUG) {
  Logger.log(`DEBUG: Preliminary match list email sent to Webmaster instead of managers for ROC on ${date}`);
} else {
  FormApp.getActiveForm().deleteAllResponses();
}  
}

/**
* Creates a PDF of the match list for a given clinic date and type.
* 
* @param {Date} date - The date of the clinic.
* @param {string} type_code - The code representing the clinic type.
* @returns {GoogleAppsScript.Drive.File} The created PDF file.
*/
function makeMatchPDF(date, type_code) {
  // PDF Creation https://developers.google.com/apps-script/samples/automations/generate-pdfs
  const pdfName = `MatchList_${type_code}_${date.toISOString().split('T')[0]}.pdf`;
  const sheet = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];

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
  const folder = DriveApp.getFoldersByName("MatchListsROC").next();

  // Not entirely sure of this is necessary or if the next file query is
  const pdfFile = folder.createFile(blob);
  //return pdfFile;

  var file = DriveApp.getFilesByName(pdfName).next();

  return file;
}