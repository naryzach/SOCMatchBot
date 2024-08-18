// Need to update SHEET ID's with new instances (link new sheet to responses first)
// Run createTriggers() after Match Tracker is updated
// Update clinic dates for proper flow of automation

// Sheet IDs
const SHEETS_ID = {
  TRACKER: "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU",  // Match tracker
  DATES: "1eztpAJ1JSWps9SiHl6pHV_6y6pWZPZ_i9tP029nw7JQ",    // Clinic dates
  MATCH: "1FW1wiYIG9LvCEgFfpgf6c9GyKljNQbwEZkoRUE3lYWg",    // Match list
  SIGN: "1FNhu-fzahjrlL0K--Mh-gSxBmDqqYWrBPG1VBOjbttQ",     // Form responses
  PEOPLE: "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"    // Contact info
};

const FORMS_ID = {
  OFFICIAL: "1FAIpQLSdeCfV1VzuK3clXnHaExcLQ88ekVoKWMMvTKql97WFr8p9WZQ",  // Main form
  MOD: "1FAIpQLSfPWCXGNPvNqVWOz2y0GYORi2_jxhtHU-vnvveXsZoUX8E54w",      // Modification form
};

const NAMES_ID = "2021179574";  // data-item-id for names list

// Signup timing (in days)
const SIGNUP_DAYS = {
  LEAD: 7,    // Open signup
  MANAGE: 3,  // Send preliminary match to manager
  CLOSE: 2    // Close signup
};

// Column indices
// Sign up sheet
const SIGN_INDEX = {
  NAME: 2,
  PTS_ALONE: 3,
  SPANISH: 4,
  SOC_POS: 5,
  FOLLOW: 6,
  CARPOOL: 7,
  COMMENTS: 8,
  DATE: 9,
  CLINIC_TYPE: 10
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
  NAMES: 15,
  MANAGERS: "A8",
  TITLE: "A1",
  DATE: "A3",
  TIME: "C3",
  ADDRESS: "A5",
  CHALK: "D12",
  PHYS: "A12",
  INTERPRET: "C8",
  SHADOW: "D8",
  VOLUNT: "C12"
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
 * Creates form submit installable triggers.
 */
/**
 * Creates form submit and time-based triggers if they don't already exist.
 */
function createTriggers() {
  const form = FormApp.getActiveForm();
  const currentTriggers = ScriptApp.getProjectTriggers();

  if (currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  ScriptApp.newTrigger("onFormSubmit")
    .forForm(form)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("updateForm")
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();

  updateStudents();
}

/**
 * Removes all triggers associated with the current project.
 */
function discontinueTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
 * Updates the form date and manages the signup process.
 * This function is triggered daily to check upcoming clinics and manage the signup process.
 */
function updateForm() {
  const form = FormApp.getActiveForm();
  const spreadsheet = SpreadsheetApp.openById(SHEETS_ID.DATES);
  const dateColumn = spreadsheet.getRange('A:A');
  dateColumn.setNumberFormat('dd-MM-yyyy');

  const today = new Date();
  const checkingDates = {
    signUp: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.LEAD),
    close: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.CLOSE),
    manage: new Date(today.getFullYear(), today.getMonth(), today.getDate() + SIGNUP_DAYS.MANAGE)
  };

  const lastRow = spreadsheet.getSheets()[0]
    .getRange(1, 1)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();

  for (let i = 2; i <= lastRow; i++) {
    const clinicDate = new Date(spreadsheet.getRange(`A${i}`).getValue());
    const timeZone = `GMT-${clinicDate.getTimezoneOffset() / 60}`; // Note: This won't work east of Prime Meridian
    const dateString = Utilities.formatDate(clinicDate, timeZone, 'EEEE, MMMM dd, YYYY');

    const clinicInfo = getClinicInfo(clinicDate, spreadsheet.getRange(`B${i}`).getValue());

    const links = {
      form: `https://docs.google.com/forms/d/e/${FORMS_ID.OFFICIAL}/viewform?usp=sf_link`,
      formMod: `https://docs.google.com/forms/d/e/${FORMS_ID.MOD}/viewform?usp=sf_link`,
      date: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.DATES}/edit?usp=sharing`,
      track: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`,
      match: `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`
    };

    if (clinicDate.valueOf() === checkingDates.signUp.valueOf()) {
      handleSignUp(form, dateString, clinicDate, timeZone, clinicInfo, links);
    } else if (clinicDate.valueOf() === checkingDates.manage.valueOf()) {
      handlePreliminaryMatch(form, dateString, clinicDate, clinicInfo);
    } else if (clinicDate.valueOf() === checkingDates.close.valueOf()) {
      handleFinalMatch(dateString, clinicDate, clinicInfo);
    }
  }
}

/**
 * Retrieves clinic information based on the date and type code.
 * @param {Date} date - The date of the clinic.
 * @param {string} typeCode - A code representing the clinic type (Y, SS, or F).
 * @returns {Object} An object containing clinic information (type, address, rooms, time, typeCode).
 */
function getClinicInfo(date, typeCode) {
  const clinicTypes = {
    Y: { type: "Yerington", address: "South Lyon Physicians Clinic, 213 S Whitacre St., Yerington,NV", rooms: 6 },
    SS: { type: "Silver Springs", address: "3595 Hwy 50, Suite 3 (In Lahontan Medical Complex), Silver Springs, NV", rooms: 5 },
    F: { type: "Fallon", address: "485 West B St Suite 101, Fallon, NV", rooms: 3 }
  };

  const clinicInfo = clinicTypes[typeCode] || { type: "Unknown", address: "Unknown", rooms: 0 };
  clinicInfo.time = date.getDay() === 0 ? "9am - 3pm" : date.getDay() === 6 ? "9am - 1pm" : "Unknown";
  clinicInfo.typeCode = typeCode;

  return clinicInfo;
}

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
    to: GET_INFO("ClassLists", "email"),
    subject: `Sign up for ROC on ${dateString}`,
    replyTo: GET_INFO("ROCManager", "email"),
    htmlBody: emailHtml,
    name: "ROC Scheduling Assistant"
  });
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
    to: GET_INFO("ClassLists", "email"),
    subject: `Match list for ROC on ${dateString}`,
    replyTo: GET_INFO("ROCManager", "email"),
    htmlBody: emailHtml,
    attachments: [file.getAs(MimeType.PDF)],
    name: "ROC Scheduling Assistant"
  });
}

/**
 * Updates the match list for a clinic based on sign-ups and participant scores.
 * This function handles the preliminary matching process, including:
 * - Gathering sign-ups for the current clinic date
 * - Calculating match scores for each participant
 * - Generating a sorted match list
 * - Updating the match list spreadsheet
 * - Sending a preliminary match list email to managers
 *
 * @param {Date} date - The date of the clinic
 * @param {string} type - The type code of the clinic (e.g., "Y" for Yerington)
 * @param {number} num_rooms - The number of available rooms for the clinic
 * @param {string} address - The address of the clinic
 */
function updateMatchList(date, type, num_rooms, address) {
  const sheetMatch = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];
  const sheetsTrack = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const sheetSign = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];

  // Gather names of signups for current dated clinic
  const lastRow = sheetSign.getLastRow();
  const signDates = sheetSign.getRange(2, SIGN_INDEX.DATE, lastRow - 1, 1).getValues();
  const signNames = sheetSign.getRange(2, SIGN_INDEX.NAME, lastRow - 1, 1).getValues();
  
  const largeNameList = signDates.reduce((acc, dateRow, index) => {
    if (dateRow[0].valueOf() === date.valueOf()) {
      acc.push(signNames[index][0]);
    }
    return acc;
  }, []);

  const namesWithScores = {};

  // Generate match list
  largeNameList.forEach(name => {
    const nameRowIndex = signNames.findIndex(row => row[0] === name) + 2;
    const nameArr = findCellByName(name);
    
    if (nameArr[0] === -1) {
      Logger.log(`Name error: ${name}`);
      if (name.endsWith("CXL")) {
        const newName = name.slice(0, -3);
        const newNameArr = findCellByName(newName);
        if (newNameArr[0] === -1) return;

        // Update the sign up counter if cancellation
        const cxlEarlyCell = sheetsTrack[newNameArr[0]].getRange(newNameArr[1] + 1, TRACK_INDEX.CXLEARLY);
        const cxlEarlyValue = cxlEarlyCell.getValue() || 0;
        cxlEarlyCell.setValue(cxlEarlyValue + 1);
      }
      return;
    }

    const trackSheet = sheetsTrack[nameArr[0]];
    const trackRow = nameArr[1] + 1;

    const signUps = parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.SIGNUPS).getValue()) || 0;
    const matches = parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.MATCHES).getValue()) || 0;
    const noShow = parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.NOSHOW).getValue()) || 0;
    const cxlLate = parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.CXLLATE).getValue()) || 0;
    const cxlEarly = parseInt(trackSheet.getRange(trackRow, TRACK_INDEX.CXLEARLY).getValue()) || 0;
    const lastDate = trackSheet.getRange(trackRow, TRACK_INDEX.DATE).getValue();

    const socPos = sheetSign.getRange(nameRowIndex, SIGN_INDEX.SOC_POS).getValue();
    const spanish = sheetSign.getRange(nameRowIndex, SIGN_INDEX.SPANISH).getValue();

    let matchScore = signUps - matches;

    // Add points for SOC position
    if (socPos === "Yes" && nameArr[0] <= 1) {
      matchScore *= 2;
    }

    // Add points for Spanish
    if (spanish === "Yes") {
      matchScore += 35;
    }

    // Add points for years of experience
    matchScore += [0, 50, 500, 1000, 0, 0][nameArr[0]] || 0;

    if (!lastDate) {
      matchScore += 25;
    } else {
      const daysSince = (new Date() - new Date(lastDate)) / (1000 * 60 * 60 * 24);
      matchScore += daysSince / 365;
    }

    matchScore -= (noShow * 3) + (cxlLate * 2) + cxlEarly;

    namesWithScores[name] = matchScore;
  });

  Logger.log("Match scores:", namesWithScores);

  // Generate sorted match list based on scores
  const sortedMatchList = Object.entries(namesWithScores)
    .sort(([, scoreA], [, scoreB]) => scoreB - scoreA)
    .map(([name]) => name)
    .slice(0, num_rooms * 2);

  Logger.log("Sorted match list:", sortedMatchList);

  // Clear Match List Sheet
  const clearRange = sheet_match.getRange(MATCH_INDEX.NAMES, 1, 25, 3);
  clearRange.clearContent().setBorder(false, false, false, false, false, false);

  // Clear specific fields
  const fieldsToClean = [MATCH_INDEX.CHALK, MATCH_INDEX.INTERPRET, MATCH_INDEX.SHADOW, MATCH_INDEX.PHYS, MATCH_INDEX.VOLUNT];
  fieldsToClean.forEach(field => sheet_match.getRange(field).clearContent());

  // Update Match List Sheet header
  const clinicInfo = {
    time: date.getDay() === 0 ? "9AM - 3PM" : date.getDay() === 6 ? "9AM - 1PM" : "Unknown",
    title: {
      "Y": "Yerington ROC Clinic",
      "SS": "Silver Springs ROC Clinic",
      "F": "Fallon ROC Clinic"
    }[type] || "Unknown"
  };

  if (clinicInfo.time === "Unknown") {
    Logger.log("Issue with time extraction");
  }
  if (clinicInfo.title === "Unknown") {
    Logger.log("Problem with clinic type");
  }

  sheet_match.getRange(MATCH_INDEX.TITLE).setValue(clinicInfo.title);
  sheet_match.getRange(MATCH_INDEX.DATE).setValue(date);
  sheet_match.getRange(MATCH_INDEX.TIME).setValue(clinicInfo.time);
  sheet_match.getRange(MATCH_INDEX.ADDRESS).setValue(address);

  // Update Match List Sheet
  let firstName, lastName, nameRowIndex;
  const actuallyMatched = [];
  const rollOverProviders = [];

  Logger.log(`Number of rooms: ${num_rooms}`);
  Logger.log(`Number of providers: ${matchList.length}`);

  num_rooms -= 1; // DIME takes a room space
  const numSlots = Math.min(matchList.length, num_rooms);

  Logger.log(`Number of slots: ${numSlots}`);

  // Fill rooms with people who can see patients alone
  for (let i = 0; i < numSlots; i++) {
    // Find the index of the name on the sign-up sheet
    nameRowIndex = 2;
    for (let j = 0; j < lastRow - 1; j++) {
      if (date.valueOf() == sign_dates[j][0].valueOf() && sign_names[j][0] == matchList[i]) {
        nameRowIndex += j; // List index offset from sheet
        break;
      }
    }
    const ptsAlone = sheet_sign.getRange(nameRowIndex, SIGN_INDEX.PTS_ALONE).getValue();

    if (ptsAlone === "Yes") {
      actuallyMatched.push(matchList[i]);
      const nameArr = findCellByName(matchList[i]);
      firstName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
      lastName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();
      
      // Update match list sheet with provider information
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 1).setValue(`Room ${i + 1}`);
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
      sheet_match.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);
    } else {
      rollOverProviders.push(matchList.splice(i, 1)[0]);
      i--; // Adjust index since we removed an item
    }
    if (matchList.length <= (i+1)) {num_slots = matchList.length; break;}
  }

  Logger.log(`Roll over providers: ${rollOverProviders}`);

  // Fill the second room spot
  const matchListP2 = rollOverProviders.concat(matchList.slice(numSlots));
  const numSlots2 = Math.min(matchListP2.length, numSlots);

  Logger.log(`Number of slots (for 2nd pass): ${numSlots2}`);

  // Add second provider to each room
  for (let i = 0; i < numSlots2; i++) {
    actuallyMatched.push(matchListP2[i]);
    const nameArr = findCellByName(matchListP2[i]);
    firstName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
    lastName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();

    const prevName = sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${prevName}\n${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
  }

  Logger.log(`Match list part 2: ${matchListP2}`);

  // Add post-bac spaces
  for (let i = 0; i < numSlots + 1; i++) { // Add room back for DIME
    const prevName = sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).getValue();
    sheet_match.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${prevName}\nPost-bac: `);
  }

  // Add DIME slot
  sheet_match.getRange(numSlots + MATCH_INDEX.NAMES, 1).setValue("DIME Providers");
  sheet_match.getRange(numSlots + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange(numSlots + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Update match stats and gather sign-up information
  let managerEmailBody = "";
  const signUpInfo = {};

  for (const name of actuallyMatched) {
    const nameArr = findCellByName(name);
    const trackSheet = sheets_track[nameArr[0]];
    const row = nameArr[1] + 1;

    // Update match count and date
    let matches = trackSheet.getRange(row, TRACK_INDEX.MATCHES).getValue() || 0;
    trackSheet.getRange(row, TRACK_INDEX.MATCHES).setValue(matches + 1);
    trackSheet.getRange(row, TRACK_INDEX.DATE).setValue(date);

    // Gather sign-up information
    const nameRowIndex = sign_names.findIndex(n => n[0] === name && sign_dates[n[0]].valueOf() === date.valueOf()) + 2;
    if (nameRowIndex > 1) {
      signUpInfo[name] = {
        spanish: sheet_sign.getRange(nameRowIndex, SIGN_INDEX.SPANISH).getValue(),
        follow: sheet_sign.getRange(nameRowIndex, SIGN_INDEX.FOLLOW).getValue(),
        carpool: sheet_sign.getRange(nameRowIndex, SIGN_INDEX.CARPOOL).getValue(),
        comments: sheet_sign.getRange(nameRowIndex, SIGN_INDEX.COMMENTS).getValue()
      };

      managerEmailBody += `${name} -- Speaks Spanish: ${signUpInfo[name].spanish}; Can have followers: ${signUpInfo[name].follow}; Carpool status: ${signUpInfo[name].carpool}; Comments: ${signUpInfo[name].comments}\n`;
    }
  }

  // Send email with the preliminary match list for Managers to update
  const htmlBody = HtmlService.createTemplateFromFile('MatchPrelimEmail');
  const timeZone = `GMT-${date.getTimezoneOffset() / 60}`; // Note: This won't work east of Prime Meridian
  const dateString = Utilities.formatDate(date, timeZone, 'EEEE, MMMM dd, YYYY');
  const linkMatch = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.MATCH}/edit?usp=sharing`;
  const linkTrack = `https://docs.google.com/spreadsheets/d/${SHEETS_ID.TRACKER}/edit?usp=sharing`;

  htmlBody.type = clinic_title;
  htmlBody.date = dateString;
  htmlBody.time = clinic_time;
  htmlBody.link = linkMatch;
  htmlBody.link_track = linkTrack;
  htmlBody.sign_up_notes = managerEmailBody;

  MailApp.sendEmail({
    to: `${GET_INFO("ROCManager", "email")},${GET_INFO("DIMEManager", "email")},${GET_INFO("LayCouns", "email")}`,
    subject: "ROC Match List (Prelim) and Notes from Sign-ups",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: htmlBody.evaluate().getContent(),
    name: "ROC Scheduling Assistant"
  });

  FormApp.getActiveForm().deleteAllResponses();
}

/**
 * Creates a PDF of the match list for a given clinic date and type.
 * 
 * @param {Date} date - The date of the clinic.
 * @param {string} type_code - The code representing the clinic type.
 * @returns {GoogleAppsScript.Drive.File} The created PDF file.
 */
function makeMatchPDF(date, type_code) {
  const pdfName = `MatchList_${type_code}_${date.toISOString().split('T')[0]}.pdf`;
  const sheet = SpreadsheetApp.openById(SHEETS_ID.MATCH).getSheets()[0];

  const exportOptions = {
    format: 'pdf',
    size: '7',
    fzr: 'true',
    portrait: 'true',
    fitw: 'true',
    gridlines: 'false',
    printtitle: 'false',
    top_margin: '0.25',
    bottom_margin: '0.25',
    left_margin: '0.25',
    right_margin: '0.25',
    sheetnames: 'false',
    pagenum: 'UNDEFINED',
    attachment: 'true',
    gid: sheet.getSheetId(),
    r1: 0,
    c1: 0,
    r2: 30,
    c2: 4
  };

  const url = `https://docs.google.com/spreadsheets/${SHEETS_ID.MATCH}/export?${Object.entries(exportOptions).map(([k, v]) => `${k}=${v}`).join('&')}`;

  const params = { 
    method: "GET", 
    headers: { "authorization": `Bearer ${ScriptApp.getOAuthToken()}` } 
  };

  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName);
  const folder = DriveApp.getFoldersByName("MatchListsROC").next();
  folder.createFile(blob);

  return DriveApp.getFilesByName(pdfName).next();
}

/**
 * Handles form submission for ROC clinic sign-ups.
 * 
 * This function is triggered when a form is submitted. It performs the following tasks:
 * 1. Logs the submitted name for error checking.
 * 2. Prevents resubmission for the same person and date.
 * 3. Updates the sign-up sheet with the new entry.
 * 4. Increments the sign-up counter in the tracker sheet for the participant.
 * 
 * @param {Object} e - The form submit event object containing response data.
 */
function onFormSubmit(e) {
  const form = FormApp.getActiveForm();
  const sheet = SpreadsheetApp.openById(SHEETS_ID.SIGN).getSheets()[0];
  const sheetsTracker = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  
  const formResponse = e.response;
  const name = formResponse.getItemResponses()[0].getResponse();
  Logger.log(name); // Log name for error checking

  const [date, clinicTypeCode] = form.getDescription().split(";");
  const lastRow = sheet.getLastRow();
  
  // Prevent resubmission 
  const usedNames = sheet.getRange(2, SIGN_INDEX.NAME, lastRow - 1).getValues();
  const usedDates = sheet.getRange(2, SIGN_INDEX.DATE, lastRow - 1).getValues();
  
  for (let i = 0; i < lastRow - 2; i++) {
    if (name === usedNames[i][0] && new Date(date).valueOf() === usedDates[i][0].valueOf()) {
      Logger.log(`Form resubmission for ${name}`);
      return;
    }
  }

  if (isNaN(new Date(date).valueOf())) {
    Logger.log(`Bad date for ${name}: ${date}`);
    return;
  }

  // Update sign-up sheet
  sheet.getRange(lastRow, SIGN_INDEX.DATE).setValue(date);
  sheet.getRange(lastRow, SIGN_INDEX.CLINIC_TYPE).setValue(clinicTypeCode);

  // Update tracker sheet
  const nameArr = findCellByName(name);
  const trackerCell = sheetsTracker[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.SIGNUPS);
  const currentSignups = trackerCell.getValue() || 0;
  trackerCell.setValue(currentSignups + 1);
}

/**
 * Builds a list of student names from the tracker spreadsheet.
 * 
 * This function performs the following tasks:
 * 1. Iterates through all sheets in the tracker spreadsheet.
 * 2. Extracts last names and first names from each sheet.
 * 3. Combines names with year tags (e.g., MS1, MS2) based on sheet index.
 * 4. Sorts the final list of names alphabetically.
 * 
 * @returns {string[]} An array of formatted student names (e.g., "Last, First (MS1)").
 */
function buildNameList() {
  const sheets = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const studentNames = [];

  sheets.forEach((sheet, index) => {
    const yearTag = getYearTag(index);
    if (!yearTag) return;

    const lastNames = sheet.getRange(2, TRACK_INDEX.LASTNAME, sheet.getLastRow() - 1, 1).getValues();
    const firstNames = sheet.getRange(2, TRACK_INDEX.FIRSTNAME, sheet.getLastRow() - 1, 1).getValues();

    lastNames.forEach((lastName, rowIndex) => {
      if (lastName[0] !== "") {
        const newName = `${lastName[0]}, ${firstNames[rowIndex][0]} (${yearTag})`;
        if (studentNames.includes(newName)) {
          Logger.log(`Duplicate: ${newName}`);
        } else {
          studentNames.push(newName);
        }
      }
    });
  });

  return studentNames.sort();
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
  let positionIndex;

  switch(position) {
    case "CEO": positionIndex = PEOPLE_INDEX.CEO; break;
    case "COO": positionIndex = PEOPLE_INDEX.COO; break;
    case "Webmaster": positionIndex = PEOPLE_INDEX.WEBMASTER; break;
    case "GenPedManager": positionIndex = PEOPLE_INDEX.GEN_PED; break;
    case "WomenManager": positionIndex = PEOPLE_INDEX.WOMEN; break;
    case "GeriDermManager": positionIndex = PEOPLE_INDEX.GERI_DERM; break;
    case "DIMEManager": positionIndex = PEOPLE_INDEX.DIME; break;
    case "ROCManager": positionIndex = PEOPLE_INDEX.ROC; break;
    case "SMManager": positionIndex = PEOPLE_INDEX.SM; break;
    case "LayCouns": positionIndex = PEOPLE_INDEX.LAY; break;
    case "ClassLists": positionIndex = PEOPLE_INDEX.CLASS; break;
    default: positionIndex = -1;
  }

  if (positionIndex === -1) {
    return info.toLowerCase() === "name" ? "Name Not Found" : "Email Not Found";
  }

  const name = sheet.getRange(positionIndex, 2).getValue();
  const email = sheet.getRange(positionIndex, 3).getValue();

  switch(info.toLowerCase()) {
    case "email":
      return email;
    case "name":
      return name;
    default:
      return "Bad Lookup";
  }
}