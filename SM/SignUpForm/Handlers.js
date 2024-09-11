/**
 * Street Medicine (SM) Sign-Up Form Handlers
 * 
 * This script contains handler functions for managing the sign-up process,
 * preliminary and final matching, and email communications for the Street Medicine program.
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
 * - Ensure this file is used in conjunction with the main SM scheduling script
 * - Update email recipients and content as needed for your specific SM program
 * - Verify that all referenced spreadsheet IDs and ranges are correct
 */

/**
 * Sends sign-up email and updates the form for an upcoming clinic.
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form to update.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {string} timeZone - Time zone string for date formatting.
 * @param {Object} links - Object containing various relevant URLs.
 */
function handleSignUp(form, dateString, clinicDate, timeZone, links) {
  updateStudents();
  form.setTitle(`Street Medicine Clinic Sign Up -- ${dateString} from 8am - 12pm`);
  form.setDescription(Utilities.formatDate(clinicDate, timeZone, 'MM/dd/YYYY'));
  form.setAcceptingResponses(true);

  const formCloseDate = new Date();
  formCloseDate.setDate(formCloseDate.getDate() + (SIGNUP_DAYS.LEAD - SIGNUP_DAYS.MANAGE));

  const htmlBody = HtmlService.createTemplateFromFile('SignUpEmail');
  htmlBody.date = dateString;
  htmlBody.close_date = Utilities.formatDate(formCloseDate, timeZone, 'EEEE, MMMM dd, YYYY');
  htmlBody.link = links.form;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");

  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("ClassLists", "email"),
    subject: `Sign up for Street Medicine Clinic on ${dateString}`,
    replyTo: GET_INFO("SMManager", "email"),
    htmlBody: htmlBody.evaluate().getContent(),
    name: "Street Medicine Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Sign-up email sent to Webmaster instead of class lists for clinic on ${dateString}`);
  }
}

/**
 * Sends preliminary match list and closes the sign-up form.
 * @param {GoogleAppsScript.Forms.Form} form - The Google Form to update.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The spreadsheet containing clinic information.
 */
function handlePreliminaryMatch(form, clinicDate, spreadsheet) {
  form.setTitle("Sign Ups Closed.");
  form.setDescription("Thank you for your interest. Please check back when another clinic is closer.");
  form.setAcceptingResponses(false);

  const numRooms = parseInt(spreadsheet.getRange(DATE_INDEX.ROOMS).getValue()) || DATE_INDEX.DEFAULT_ROOMS;
  updateMatchList(clinicDate, numRooms);
}

/**
 * Sends final match list with PDF attachment to participants.
 * @param {string} dateString - Formatted date string for the clinic.
 * @param {Date} clinicDate - Date object for the clinic.
 */
function handleFinalMatch(dateString, clinicDate) {
  const file = makeMatchPDF(clinicDate);
  const htmlBody = HtmlService.createTemplateFromFile('MatchEmail');
  htmlBody.date = dateString;
  htmlBody.feedback_email = GET_INFO("Webmaster", "email");

  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("ClassLists", "email"),
    subject: `Match list for Street Medicine Clinic on ${dateString}`,
    replyTo: GET_INFO("SMManager", "email"),
    htmlBody: htmlBody.evaluate().getContent(),
    attachments: [file.getAs(MimeType.PDF)],
    name: "Street Medicine Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Final match list email sent to Webmaster instead of class lists for clinic on ${dateString}`);
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
* 2. Populates the sheet with clinic information (title, date, time).
* 3. Applies formatting to the sheet (borders, text wrapping).
*/
function setupMatchList(matchList, clinicTime, clinicInfo, date, numRooms) {
  const sheetMatch = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();

  // Clear Match List Sheet file names and remove borders
  sheetMatch.getRange(MATCH_INDEX.NAMES, 1, 25, 3).clearContent().setBorder(false, false, false, false, false, false);

  sheetMatch.getRange(MATCH_INDEX.DATE).setValue(date);
  sheetMatch.getRange(MATCH_INDEX.TIME).setValue(clinicTime);
  sheetMatch.getRange(MATCH_INDEX.TITLE).setValue(clinicInfo);

  const numSlots = Math.min(matchList.length, numRooms);

  Logger.log(`Number of providers: ${matchList.length}`);
  Logger.log(`Number of slots: ${numSlots}`);

  // Fill rooms with people who can see patients alone
  const actuallyMatched = [];
  for (let i = 0; i < numSlots; i++) {
    actuallyMatched.push(matchList[i]);
    const nameArr = findCellByName(matchList[i]);
    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1];
    const firstName = trackSheet.getRange(trackRow, TRACK_INDEX.FIRSTNAME).getValue();
    const lastName = trackSheet.getRange(trackRow, TRACK_INDEX.LASTNAME).getValue();

    const matchRow = sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3);
    matchRow.setValues([
      [`Student ${i + 1}`, 
      `${firstName} ${lastName}, ${getYearTag(nameArr[0])}`, 
      "_____________________________________________\n_____________________________________________\n_____________________________________________"]
    ]).setBorder(true, true, true, true, true, true);
  }
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
function updateMatchStats(actuallyMatched, date) {
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  // Get the last row of the sign-up sheet
  const lastRow = sheetSign.getLastRow();
  // Get the sign-up dates and names for all students
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow).getValues();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow).getValues();

  let managerEmailBody = "";
  actuallyMatched.forEach(name => {
    const nameArr = findCellByName(name);
    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1];
    
    if (!DEBUG) {
      const matchesCell = trackSheet.getRange(trackRow, TRACK_INDEX.MATCHES);
      matchesCell.setValue((matchesCell.getValue() || 0) + 1);
      trackSheet.getRange(trackRow, TRACK_INDEX.DATE).setValue(date);
      const allDates = trackSheet.getRange(trackRow, TRACK_INDEX.DATE_ALL).getValue();
      trackSheet.getRange(trackRow, TRACK_INDEX.DATE_ALL).setValue(allDates ? allDates + "," + date : date);
    } else {
      Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Matches incremented, Date set to ${date}`);
    }

    const nameRowIndex = signNames.findIndex(n => n[0] === name && signDates[signNames.indexOf(n)][0].valueOf() === date.valueOf()) + 2;
    const transport = sheetSign.getRange(nameRowIndex, SIGN_INDEX.ELECTIVE).getValue();
    const comments = sheetSign.getRange(nameRowIndex, SIGN_INDEX.COMMENTS).getValue();

    managerEmailBody += `${name} -- Reliable transport: ${transport}; Comments: ${comments}\n`;
  });

  // Send email with the preliminary match list for Managers to update
  const htmlBody = HtmlService.createTemplateFromFile('PrelimMatchEmail');
  const timeZone = "GMT-" + String(date.getTimezoneOffset() / 60); // will not work east of Prime Meridian
  const dateString = Utilities.formatDate(date, timeZone, 'EEEE, MMMM dd, YYYY');
  const linkMatch = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`;
  const linkTrack = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`;

  htmlBody.date = dateString;
  htmlBody.link = linkMatch;
  htmlBody.link_track = linkTrack;
  htmlBody.sign_up_notes = managerEmailBody;

  const emailHtml = htmlBody.evaluate().getContent();
  MailApp.sendEmail({
    to: DEBUG ? GET_INFO("Webmaster", "email") : GET_INFO("SMManager", "email"),
    subject: "Street Medicine Match List (Prelim) and Notes from Sign-ups",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: emailHtml,
    name: "SM Scheduling Assistant"
  });

  if (DEBUG) {
    Logger.log(`DEBUG: Preliminary match list email sent to Webmaster instead of managers for SM on ${date}`);
  } else {
    FormApp.getActiveForm().deleteAllResponses();
  }
}

/**
* Creates a PDF of the match list for a given date.
* @param {Date} date - The date of the clinic.
* @return {GoogleAppsScript.Drive.File} The created PDF file.
*/
function makeMatchPDF(date) {
  // Format the PDF name
  const pdfName = `MatchList_${date.toISOString().split('T')[0]}`;

  // Get the match sheet
  const sheet = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];

  // Define the range for the PDF
  const firstRow = 0, firstCol = 0, lastCol = 4, lastRow = 30;

  // Construct the URL for PDF export
  const url = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/export` +
    '?format=pdf&' +
    'size=7&' +
    'fzr=true&' +
    'portrait=true&' +
    'fitw=true&' +
    'gridlines=false&' +
    'printtitle=false&' +
    'top_margin=0.25&' +
    'bottom_margin=0.25&' +
    'left_margin=0.25&' +
    'right_margin=0.25&' +
    'sheetnames=false&' +
    'pagenum=UNDEFINED&' +
    'attachment=true&' +
    `gid=${sheet.getSheetId()}&` +
    `r1=${firstRow}&c1=${firstCol}&r2=${lastRow}&c2=${lastCol}`;

  // Set up parameters for the URL fetch
  const params = { 
    method: "GET", 
    headers: { "authorization": `Bearer ${ScriptApp.getOAuthToken()}` } 
  };

  // Fetch the PDF as a blob
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(`${pdfName}.pdf`);

  // Get the folder where PDFs are stored
  const folder = DriveApp.getFoldersByName("MatchListsSM").next();

  // Create the PDF file in the folder
  folder.createFile(blob);

  // Return the created file
  return DriveApp.getFilesByName(`${pdfName}.pdf`).next();
}