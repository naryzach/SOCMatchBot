/**
 * Rural Outreach Clinic (ROC) Sign-Up Form and Scheduling Script
 * 
 * This script manages the sign-up process, match list generation, and email communications
 * for the Rural Outreach Clinic program. It interacts with Google Sheets and Forms to automate
 * the scheduling process for rural outreach clinics.
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
 * - Update SHEET ID's and FORM ID's with new instances as needed
 * - Run createTriggers() after updating the Match Tracker or changing form IDs
 * - Ensure clinic dates in the DATES sheet are up to date for proper automation flow
 */

const DEBUG = true;

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
  ELECTIVE: 9,
  DATE: 10,
  CLINIC_TYPE: 11
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

// NOTES:
//  Code infers year based on sheet order (MS1,2,3,4,PA1,2); could update but is already pretty simple



// *** ---------------------------------- *** // 

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

    const typeCode = spreadsheet.getRange(`B${i}`).getValue();
    const clinicTypes = {
      Y: { type: "Yerington", address: "South Lyon Physicians Clinic, 213 S Whitacre St., Yerington,NV", rooms: 6 },
      SS: { type: "Silver Springs", address: "3595 Hwy 50, Suite 3 (In Lahontan Medical Complex), Silver Springs, NV", rooms: 5 },
      F: { type: "Fallon", address: "485 West B St Suite 101, Fallon, NV", rooms: 3 }
    };

    const clinicInfo = clinicTypes[typeCode] || { type: "Unknown", address: "Unknown", rooms: 0 };
    clinicInfo.time = clinicDate.getDay() === 0 ? "9am - 3pm" : clinicDate.getDay() === 6 ? "9am - 1pm" : "Unknown";
    clinicInfo.typeCode = typeCode;

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
        if (!DEBUG) {
          cxlEarlyCell.setValue(cxlEarlyValue + 1);
        } else {
          Logger.log(`DEBUG: Would update TRACKER sheet for ${newName}: CXLEARLY incremented from ${cxlEarlyValue} to ${cxlEarlyValue + 1}`);
        }
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
  const clearRange = sheetMatch.getRange(MATCH_INDEX.NAMES, 1, 25, 3);
  clearRange.clearContent().setBorder(false, false, false, false, false, false);

  // Clear specific fields
  const fieldsToClean = [MATCH_INDEX.CHALK, MATCH_INDEX.INTERPRET, MATCH_INDEX.SHADOW, MATCH_INDEX.PHYS, MATCH_INDEX.VOLUNT];
  fieldsToClean.forEach(field => sheetMatch.getRange(field).clearContent());

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

  sheetMatch.getRange(MATCH_INDEX.TITLE).setValue(clinicInfo.title);
  sheetMatch.getRange(MATCH_INDEX.DATE).setValue(date);
  sheetMatch.getRange(MATCH_INDEX.TIME).setValue(clinicInfo.time);
  sheetMatch.getRange(MATCH_INDEX.ADDRESS).setValue(address);

  // Update Match List Sheet
  let firstName, lastName, nameRowIndex;
  const actuallyMatched = [];
  const rollOverProviders = [];

  Logger.log(`Number of rooms: ${num_rooms}`);
  Logger.log(`Number of providers: ${sortedMatchList.length}`);

  num_rooms -= 1; // DIME takes a room space
  const numSlots = Math.min(sortedMatchList.length, num_rooms);

  Logger.log(`Number of slots: ${numSlots}`);

  // Fill rooms with people who can see patients alone
  for (let i = 0; i < numSlots; i++) {
    // Find the index of the name on the sign-up sheet
    nameRowIndex = 2;
    for (let j = 0; j < lastRow - 1; j++) {
      if (date.valueOf() == signDates[j][0].valueOf() && signNames[j][0] == sortedMatchList[i]) {
        nameRowIndex += j; // List index offset from sheet
        break;
      }
    }
    const ptsAlone = sheetSign.getRange(nameRowIndex, SIGN_INDEX.PTS_ALONE).getValue();

    if (ptsAlone === "Yes") {
      actuallyMatched.push(sortedMatchList[i]);
      const nameArr = findCellByName(sortedMatchList[i]);
      firstName = sheetsTrack[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
      lastName = sheetsTrack[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();
      
      // Update match list sheet with provider information
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1).setValue(`Room ${i + 1}`);
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 2).setValue(`${firstName} ${lastName}, ${getYearTag(nameArr[0])}`);
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
      sheetMatch.getRange(i + MATCH_INDEX.NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);
    } else {
      rollOverProviders.push(sortedMatchList.splice(i, 1)[0]);
      i--; // Adjust index since we removed an item
    }
    if (sortedMatchList.length <= (i+1)) {numSlots = sortedMatchList.length; break;}
  }

  Logger.log(`Roll over providers: ${rollOverProviders}`);

  // Fill the second room spot
  const matchListP2 = rollOverProviders.concat(sortedMatchList.slice(numSlots));
  const numSlots2 = Math.min(matchListP2.length, numSlots);

  Logger.log(`Number of slots (for 2nd pass): ${numSlots2}`);

  // Add second provider to each room
  for (let i = 0; i < numSlots2; i++) {
    actuallyMatched.push(matchListP2[i]);
    const nameArr = findCellByName(matchListP2[i]);
    firstName = sheetsTrack[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.FIRSTNAME).getValue();
    lastName = sheetsTrack[nameArr[0]].getRange(nameArr[1] + 1, TRACK_INDEX.LASTNAME).getValue();

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

  // Update match stats and gather sign-up information
  let managerEmailBody = "";
  const signUpInfo = {};

  for (const name of actuallyMatched) {
    const nameArr = findCellByName(name);
    const trackSheet = sheetsTrack[nameArr[0]];
    const row = nameArr[1] + 1;

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
    const nameRowIndex = signNames.findIndex(n => n[0] === name && signDates[n[0]].valueOf() === date.valueOf()) + 2;
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
    Logger.log(`DEBUG: Preliminary match list email sent to Webmaster instead of ROC Manager, DIME Manager, and Lay Counselor for ROC on ${dateString}`);
  }

  FormApp.getActiveForm().deleteAllResponses();
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
  
  if (!DEBUG) {
    trackerCell.setValue(currentSignups + 1);
  } else {
    Logger.log(`DEBUG: Would update TRACKER sheet for ${name}: Signups incremented from ${currentSignups} to ${currentSignups + 1}`);
  }
}