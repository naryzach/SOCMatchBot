// Need to update SHEET ID's with new instances (link new sheet to responses first)
// Run createTriggers() after Match Tracker is updated
// Update clinic dates for proper flow of automation

// SHEET ID's
const SHEET_TRACKER = "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU"; // match tracker
const SHEET_DATES = "1vhqF4JpN9HZwdqkMou_AQmC61mOA49mh57N2RIzpZHY"; // sheet of clinic dates
const SHEET_MATCH = "1hUJJqmnqrDD7e6n9MLtro6WIneTQl1o76atptyssig4"; // match list sheet
const SHEET_SIGN = "1zCpaz2ketqM_EnYjnTTGPijs7_6ZKSu-VwtzcHkXA1g"; // sheet associated with form
const SHEET_PEOPLE = "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"// sheet for contact info and names
const FORM_OFFICIAL = "1FAIpQLSf3361EGf494smoIhOwi8EVAgKhYR3IPqmm6SET-2RiHaLAdw" // The form to send (not the editing form)
const FORM_MOD = "1FAIpQLSfwAJj6RyqB-u4QBVclWrFi5R9FokTEHS9jB1k8PidO75xtAQ" // for the sign-up modification form
const NAMES_ID = "2021179574"; // data-item-id in HTML -- doesn't change on copy

// Form hyperparameters //
// Sign up timing
const SIGNUP_LEAD_DAYS = 5 // 5 days before clinic
const SIGNUP_MANAGE_DAYS = 3 // 3 days before clinic -- for getting room number data from managers
const SIGNUP_CLOSE_DAYS = 2 // 2 days before clinic

// Sign up sheet indexing
const SIGN_NAME_NDX = 2
const SIGN_PTS_ALONE_NDX = 3
const SIGN_ELECTIVE_NDX = 4
const SIGN_SOC_POS_NDX = 5
const SIGN_DIET_NDX = 6
const SIGN_COMMENTS_NDX = 7
const SIGN_DATE_NDX = 8
const SIGN_CLINIC_TYPE_NDX = 9

// Match tracker sheet indexing
const TRACK_LASTNAME_NDX = 1
const TRACK_FIRSTNAME_NDX = 2
const TRACK_SIGNUPS_NDX = 3
const TRACK_MATCHES_NDX = 4
const TRACK_NOSHOW_NDX = 5
const TRACK_CXLLATE_NDX = 6
const TRACK_CXLEARLY_NDX = 7
const TRACK_DATE_NDX = 8

// Match list sheet indexing
const MATCH_NAMES = 13;
const MATCH_MANGERS = "A6";
const MATCH_TITLE = "A1";
const MATCH_DATE = "A3";
const MATCH_TIME = "C3";
const MATCH_PHYS1 = "A10";
const MATCH_PHYS2 = "B10";
const MATCH_CHALK_TALK = "C10";

// Date sheet indexing
const DATE_ROOMS = "C2"
const DEFAULT_NUM_ROOMS = "10"

// People sheet indexing
const PEOPLE_CEO = 2
const PEOPLE_C0O = 3
const PEOPLE_WEBMASTER = 4
const PEOPLE_GEN_PED = 5
const PEOPLE_WOMEN = 6
const PEOPLE_GERI_DERM = 7
const PEOPLE_DIME = 8
const PEOPLE_LAY = 9
const PEOPLE_ROC = 10
const PEOPLE_SM = 11
const PEOPLE_CLASS = 12

// NOTES:
//  Code infers year based on sheet order (MS1,2,3,4,PA1,2); could update but is already pretty simple


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
  ScriptApp.newTrigger("updateFormDate").timeBased().atHour(8).everyDays(1).create();

  // Update name list
  updateStudents();
}

// Remove triggers
function discontinueTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// Change the title of the form to coincide with the date of clinic
function updateFormDate() {
  var form = FormApp.getActiveForm();
  var spreadsheet = SpreadsheetApp.openById(SHEET_DATES);

  // Format time data
  var date_column = spreadsheet.getRange('A:A');
  date_column.setNumberFormat('dd-MM-yyyy');

  // Get date,the date in 2, 3, and 5 days
  var date = new Date();
  var checkingDate = new Date(date.getFullYear(), date.getMonth(), date.getDate() + SIGNUP_LEAD_DAYS);
  var checkingDateEnd = new Date(date.getFullYear(), date.getMonth(), date.getDate() + SIGNUP_CLOSE_DAYS);
  var checkingDateManage = new Date(date.getFullYear(), date.getMonth(), date.getDate() + SIGNUP_MANAGE_DAYS);
  var checkingDateCEO = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);

  // Set variables for the loop
  var cell;
  var c_Date = new Date();
  var type = "";
  var time = "";
  var clinic_type_code = "";
  var clinic_emails = GET_INFO("Webmaster", "email");

  // Get the last row of the table
  let ss_end = spreadsheet.getSheets()[0]
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
  for (var i = 2; i <= ss_end; i++) {
    cell = spreadsheet.getRange("A" + i.toString());
    c_Date = new Date(cell.getValue().toString());

    // Get the clinic time
    switch (c_Date.getDay()) {
      case 2: 
        time = "6pm - 10pm"
        break;
      case 6:
        time = "8am - 12pm"
        break;
      default:
        time = "Unknown"
        Logger.log("Issue with time extraction");
    }

    // Get clinic type
    clinic_type_code = spreadsheet.getRange("B" + i.toString()).getValue().toString();
    switch (clinic_type_code) {
      case "W":
        type = "Women's";
        clinic_emails = GET_INFO("WomenManager", "email");
        break;
      case "GP":
        type = "Gen/Peds";
        clinic_emails = GET_INFO("GenPedManager", "email");
        break;
      case "GD":
        type = "Geri/Derm";
        clinic_emails = GET_INFO("GeriDermManager", "email");
        break;
      default:
        type = "Unknown";
        clinic_emails = GET_INFO("Webmaster", "email");
        Logger.log("Problem with clinic type");
    }

    // Format time
    var tz = "GMT-" + String(c_Date.getTimezoneOffset()/60) // will not work east of Prime Meridian
    var date_string  = Utilities.formatDate(c_Date, tz, 'EEEE, MMMM dd, YYYY');

    // Links to Google pages
    var formLink = "https://docs.google.com/forms/d/e/" + FORM_OFFICIAL + "/viewform?usp=sf_link"
    var formLinkMod = "https://docs.google.com/forms/d/e/" + FORM_MOD + "/viewform?usp=sf_link"
    var linkDate = "https://docs.google.com/spreadsheets/d/" + SHEET_DATES + "/edit?usp=sharing"
    var linkTrack = "https://docs.google.com/spreadsheets/d/" + SHEET_TRACKER + "/edit?usp=sharing"
    var linkMatch = "https://docs.google.com/spreadsheets/d/" + SHEET_MATCH + "/edit?usp=sharing"

    // If 5 days out, update the Form
    if (c_Date.valueOf() == checkingDate.valueOf()) {
      // Update Form information
      updateStudents();
      form.setTitle(type + " Clinic Sign Up -- " + date_string + " from " + time);
      form.setDescription(Utilities.formatDate(c_Date, tz, 'MM/dd/YYYY') + ";" + clinic_type_code);
      form.setAcceptingResponses(true);
      var formCloseDate = date;
      formCloseDate.setDate(formCloseDate.getDate() + (SIGNUP_LEAD_DAYS-SIGNUP_CLOSE_DAYS));

      // Send email prompting sign ups from HTML format
      var html_body = HtmlService.createTemplateFromFile('SignUpEmail');  
      html_body.type = type;
      html_body.date = date_string;
      html_body.close_date = Utilities.formatDate(formCloseDate, tz, 'EEEE, MMMM dd, YYYY');
      html_body.time = time;
      html_body.link = formLink;
      html_body.feedback_email = GET_INFO("Webmaster", "email");
      var email_html = html_body.evaluate().getContent();
      MailApp.sendEmail({
        to: GET_INFO("ClassLists", "email"),
        subject:  "Sign up for SOC on " + date_string,
        replyTo: clinic_emails,
        htmlBody: email_html,
        name: "SOC Scheduling Assistant"
      });
    
    // if 3 days out, ask for number of rooms
    } else if (c_Date.valueOf() == checkingDateManage.valueOf()) {
      // Set number of rooms to 10 by defult
      spreadsheet.getRange(DATE_ROOMS).setValue(DEFAULT_NUM_ROOMS);

      // Send email prompting managers to fill in the number of rooms needed
      var html_body = HtmlService.createTemplateFromFile('RoomNumEmail');  
      html_body.type = type;
      html_body.date = date_string;
      html_body.cell_ndx = DATE_ROOMS;
      html_body.time = time;
      html_body.link_date = linkDate;
      html_body.link_match = linkMatch;
      html_body.link_match_track = linkTrack;
      html_body.link_form_mod = formLinkMod;
      var email_html = html_body.evaluate().getContent();
      MailApp.sendEmail({
        to: clinic_emails,
        subject:  "SOC Clinic Room Availiblity Form",
        replyTo: GET_INFO("Webmaster", "email"),
        htmlBody: email_html,
        name: "SOC Scheduling Assistant"
      });
    
    // if 2 days out, update form
    } else if (c_Date.valueOf() == checkingDateEnd.valueOf()) {
      // Update form
      form.setTitle("Sign Ups Closed.");
      form.setDescription("Thank you for your interest. Please check back when another clinic is closer.");
      //form.setDescription(Utilities.formatDate(c_Date, tz, 'MM/dd/YYYY'));
      form.setAcceptingResponses(false);
      var num_rooms = parseInt(spreadsheet.getRange(DATE_ROOMS).getValue().toString());
      updateMatchList(checkingDateEnd, clinic_type_code, num_rooms);

      // Set tmp data in dates sheet for later use
      spreadsheet.getRange("X1").setValue(type);
      spreadsheet.getRange("X2").setValue(date_string);
      spreadsheet.getRange("X3").setValue(time);
      spreadsheet.getRange("X4").setValue(clinic_emails);

      // Delay the sending the match list until noon (set up a new trigger)
      var d_today = new Date();
      var d_year = d_today.getFullYear();
      var d_month = d_today.getMonth();
      var d_day = d_today.getDate();
      var d_functionName = 'matchListDelay';
      ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === d_functionName) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      ScriptApp.newTrigger(d_functionName)
        .timeBased()
        .at(new Date(d_year, d_month, d_day, 17, 0))
        .create();
    
    /*
    // if 1 day out, send the email to CEO and COO with the PDF
    } else if (c_Date.valueOf() == checkingDateCEO.valueOf()) {
      var file = makeMatchPDF(c_Date); // make the PDF of the match list
      
      MailApp.sendEmail({
        to: GET_INFO("CEO", "email") + ", " + GET_INFO("COO", "email"),
        subject:  "Final Match List for SOC on " + date_string,
        replyTo: GET_INFO("Webmaster", "email"),
        htmlBody: "Attached is the PDF for the upcoming SOC.<br><br>Best,<br>The SOC Scheduler",
        attachments: [file.getAs(MimeType.PDF)],
        name: "SOC Scheduling Assistant"
      });

    // Save final version of Match List on the day of clinic
    } else if (c_Date.valueOf() == date.valueOf()) {
      makeMatchPDF(c_Date)
    */
    }
  }
}

// Handler function for the delayed match list disemination
function matchListDelay() {
  // Grab tmp data (from date sheet)
  var spreadsheet = SpreadsheetApp.openById(SHEET_DATES);
  var type = spreadsheet.getRange("X1").getValue().toString();
  var raw_date = new Date(spreadsheet.getRange("X2").getValue().toString());
  var tz = "GMT-" + String(raw_date.getTimezoneOffset()/60)
  var date_string = Utilities.formatDate(raw_date, tz, 'EEEE, MMMM dd, YYYY')
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
    to: GET_INFO("ClassLists", "email"),
    subject:  "Match list for SOC on " + date_string,
    replyTo: clinic_emails,
    htmlBody: email_html,
    attachments: [file.getAs(MimeType.PDF)],
    name: "SOC Scheduling Assistant"
  });
}

// Create match list and apply changes to Sheets file
function updateMatchList(date, type, num_rooms) {
  var sheet_match = SpreadsheetApp.openById(SHEET_MATCH).getSheets()[0];
  var sheets_track = SpreadsheetApp.openById(SHEET_TRACKER).getSheets();
  var sheet_sign = SpreadsheetApp.openById(SHEET_SIGN).getSheets()[0];

  var largeNameList = [];

  // Gather names of signups for current dated clinic
  let lastRow = sheet_sign
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();
  var sign_dates = sheet_sign.getRange(2, SIGN_DATE_NDX, lastRow).getValues();
  var sign_names = sheet_sign.getRange(2, SIGN_NAME_NDX, lastRow).getValues();
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
        var tmp = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_CXLEARLY_NDX).getValue();
        if (tmp == "") { 
          tmp = 0;
        }
        sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_CXLEARLY_NDX).setValue(parseInt(tmp) + 1);
      }

      // Do not try to match that name
      continue;
    }
    signUps = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_SIGNUPS_NDX).getValue());
    matches = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_MATCHES_NDX).getValue());
    noShow = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_NOSHOW_NDX).getValue());
    cxlLate = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_CXLLATE_NDX).getValue());
    cxlEarly = parseInt(sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_CXLEARLY_NDX).getValue());
    lastDate = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_DATE_NDX).getValue();

    // Grab form submission data
    fourthYrElect = sheet_sign.getRange(name_row_ndx, SIGN_ELECTIVE_NDX).getValue();
    socPos = sheet_sign.getRange(name_row_ndx, SIGN_SOC_POS_NDX).getValue();

    // Caluclate match score
    matchScore = signUps - matches;

    // Elective and SOC position additions
    if (socPos == "Yes" && (nameArr[0] == 0 || nameArr[0] == 1)) { //MS1/2s -- second check is unneccesary
      //matchScore += 100; // rank SOC members in a hierarchy
      matchScore *= 2; // Only slightly bias SOC members rather than rank in a hierarchy
    }
    if (fourthYrElect == "Yes" && nameArr[0] == 3) { // MS4s
      matchScore += 500;
    }

    // Add points based on seniority
    switch (nameArr[0]) {
      case 0:
        break;
      case 1:
        matchScore += 50; //second year
        break;
      case 2:
        matchScore += 500; //third year
        break;
      case 3:
        matchScore += 1000; //fourth year
        break;
      case 4:
        break;
      case 5:
        break;
    }

    // Never been matched addition
    var daysSince = 0;
    if (lastDate == "") {
      matchScore += 25;
    // Add fractional points to help sort by last match
    } else {
      daysSince = (new Date - new Date(lastDate)) / (1000*60*60*24);
      matchScore += daysSince/365;
    }

    // Cancellation penalty
    if (!isNaN(noShow)) {
      matchScore -= noShow * 3;
    }
    if (!isNaN(cxlLate)) {
      matchScore -= cxlLate * 2;
    }
    if (!isNaN(cxlEarly)) {
      matchScore -= cxlEarly;
    }

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
    sheet_match.getRange(i + MATCH_NAMES, 1).setValue("");
    sheet_match.getRange(i + MATCH_NAMES, 2).setValue("");
    sheet_match.getRange(i + MATCH_NAMES, 3).setValue("");
    sheet_match.getRange(i + MATCH_NAMES, 1, 1, 3).setBorder(false, false, false, false, false, false);
  }
  // Clear physicians and chalk talk
  sheet_match.getRange("A10").setValue("");
  sheet_match.getRange("B10").setValue("");
  sheet_match.getRange("C10").setValue("");

  // Update Match List Sheet header
  var dayOfWeek = date.getDay();
  var clinic_time = "";
  switch (dayOfWeek) {
    case 2: 
      clinic_time = "6PM - 10PM";
      break;
    case 6:
      clinic_time = "8AM - 12PM";
      break;
    default:
      clinic_time = "Unknown";
      Logger.log("Issue with time extraction");
  }

  var clinic_title = "";
  var manager_names = "";
  var manager_emails = "";
  switch (type) {
    case "W":
      clinic_title = "Women's Clinic";
      manager_names = GET_INFO("WomenManager", "name");
      manager_emails = GET_INFO("WomenManager", "email");
      break;
    case "GP":
      clinic_title = "General & Pediatric Clinic";
      manager_names = GET_INFO("GenPedManager", "name");
      manager_emails = GET_INFO("GenPedManager", "email");
      break;
    case "GD":
      clinic_title = "Geriatrics & Dermatology Clinic";
      manager_names = GET_INFO("GeriDermManager", "name");
      manager_emails = GET_INFO("GeriDermManager", "email");
      break;
    default:
      clinic_title = "Unknown";
      manager_names = "Error";
      manager_emails = GET_INFO("Webmaster", "email")
      Logger.log("Problem with clinic type");
  }
  sheet_match.getRange(MATCH_TITLE).setValue(clinic_title);
  sheet_match.getRange(MATCH_DATE).setValue(date);
  sheet_match.getRange(MATCH_TIME).setValue(clinic_time);
  sheet_match.getRange(MATCH_MANGERS).setValue(manager_names);

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
    ptsAlone = sheet_sign.getRange(name_row_ndx, SIGN_PTS_ALONE_NDX).getValue();

    if (ptsAlone == "Yes") {
      actuallyMatched.push(matchList[i]);
      nameArr = findCellByName(matchList[i]);
      firstName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_FIRSTNAME_NDX).getValue();
      lastName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_LASTNAME_NDX).getValue();
      sheet_match.getRange(i + MATCH_NAMES, 1).setValue("Room " + (i + 1).toString());
      sheet_match.getRange(i + MATCH_NAMES, 2).setValue(firstName + " " + lastName + ", " + getYearTag(nameArr[0]));
      sheet_match.getRange(i + MATCH_NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
      sheet_match.getRange(i + MATCH_NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);
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
    firstName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_FIRSTNAME_NDX).getValue();
    lastName = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_LASTNAME_NDX).getValue();

    prev_name = sheet_match.getRange(i + MATCH_NAMES, 2).getValue();
    sheet_match.getRange(i + MATCH_NAMES, 2).setValue(prev_name + "\n" + firstName + " " + lastName + ", " + getYearTag(nameArr[0]));
  }

  Logger.log("Match list part 2")
  Logger.log(matchListP2)

  // Add volunteer spaces
  for (i = 0; i < num_slots; i++) {
    prev_name = sheet_match.getRange(i + MATCH_NAMES, 2).getValue();
    sheet_match.getRange(i + MATCH_NAMES, 2).setValue(prev_name + "\nVolunteer: ");
  }

  // Add DIME Manager slot
  sheet_match.getRange(num_slots + MATCH_NAMES, 1).setValue("DIME Managers");
  sheet_match.getRange(num_slots + MATCH_NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange((num_slots+1) + MATCH_NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Add DIME Provider slot
  sheet_match.getRange((num_slots+1) + MATCH_NAMES, 1).setValue("DIME Providers");
  sheet_match.getRange((num_slots+1) + MATCH_NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange(num_slots + MATCH_NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // Add lay counselor slot
  sheet_match.getRange((num_slots+2) + MATCH_NAMES, 1).setValue("Lay Counselors");
  sheet_match.getRange((num_slots+2) + MATCH_NAMES, 2).setValue(GET_INFO("LayCouns", "name"));
  sheet_match.getRange((num_slots+2) + MATCH_NAMES, 3).setValue("_____________________________________________\n_____________________________________________\n_____________________________________________");
  sheet_match.getRange((num_slots+2) + MATCH_NAMES, 1, 1, 3).setBorder(true, true, true, true, true, true);

  // ** Mutable changes after here ** // 

  // Update match stats
  var comments = "";
  var diet_restrict = "";
  var manager_email_body = "";
  for (name in actuallyMatched) {
    nameArr = findCellByName(actuallyMatched[name])
    var tmp = sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_MATCHES_NDX).getValue();
    if (tmp == "") tmp = 0;
    sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_MATCHES_NDX).setValue(parseInt(tmp) + 1);
    sheets_track[nameArr[0]].getRange(nameArr[1] + 1, TRACK_DATE_NDX).setValue(date);

    // Account for dietary restrictions and comments
    name_row_ndx = 2;
    for(var j = 0; j < lastRow-1; j++) {
      if (date.valueOf() == sign_dates[j][0].valueOf() && sign_names[j][0] == actuallyMatched[name])
        name_row_ndx += j; // List index offset from sheet
    }
    diet_restrict = sheet_sign.getRange(name_row_ndx, SIGN_DIET_NDX).getValue();
    comments = sheet_sign.getRange(name_row_ndx, SIGN_COMMENTS_NDX).getValue(); 

    if ((diet_restrict != "None" && diet_restrict != "") || comments != "") {
      manager_email_body += actuallyMatched[name] + " -- Dietary restrictions: " + diet_restrict + "; Comments: " + comments + "\n"; 
    }
  }

  if (manager_email_body == "") {
    manager_email_body = "No comments or dietary restictions noted by matched students.";
  }

  // Send email prompting managers to fill in the number of rooms needed
  var html_body = HtmlService.createTemplateFromFile('PrelimMatchEmail');
  var linkMatch = "https://docs.google.com/spreadsheets/d/" + SHEET_MATCH + "/edit?usp=sharing";
  html_body.link_match = linkMatch;
  html_body.sign_up_notes = manager_email_body;
  var email_html = html_body.evaluate().getContent();
  MailApp.sendEmail({
    to: manager_emails + "," + GET_INFO("DIMEManager", "email") + "," + GET_INFO("LayCouns", "email"),
    subject:  "Notes from SOC sign up",
    replyTo: GET_INFO("Webmaster", "email"),
    htmlBody: email_html,
    name: "SOC Scheduling Assistant"
  });

  FormApp.getActiveForm().deleteAllResponses();
}

function makeMatchPDF(date) {
  // PDF Creation https://developers.google.com/apps-script/samples/automations/generate-pdfs
  pdfName = "MatchList" + (date).toISOString().split('T')[0] + ".pdf"
  var sheet = SpreadsheetApp.openById(SHEET_MATCH).getSheets()[0];

  const fr = 0, fc = 0, lc = 4, lr = 30;
  const url = "https://docs.google.com/spreadsheets/d/" + SHEET_MATCH + "/export" +
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

// A function that is called by the form submit
// trigger. The parameter e contains information
// submitted by the user.
function onFormSubmit(e) {
  var form = FormApp.getActiveForm();
  var sheet = SpreadsheetApp.openById(SHEET_SIGN).getSheets()[0];
  var sheets_tracker = SpreadsheetApp.openById(SHEET_TRACKER).getSheets();
  
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
  var usedNames = sheet.getRange(2, SIGN_NAME_NDX, lastRow-1).getValues();
  var usedDates = sheet.getRange(2, SIGN_DATE_NDX, lastRow-1).getValues();
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
  sheet.getRange(lastRow, SIGN_DATE_NDX).setValue(date);
  sheet.getRange(lastRow, SIGN_CLINIC_TYPE_NDX).setValue(clinic_type_code);

  // Update the sign up counter
  var tmp = sheets_tracker[nameArr[0]].getRange(nameArr[1] + 1, TRACK_SIGNUPS_NDX).getValue();
  if (tmp == "") { 
    tmp = 0;
  }
  sheets_tracker[nameArr[0]].getRange(nameArr[1] + 1, TRACK_SIGNUPS_NDX).setValue(parseInt(tmp) + 1);
}

// Send me the match list
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

// Build the list of names from the sheet
function buildNameList () {
  var sheets = SpreadsheetApp.openById(SHEET_TRACKER).getSheets();
  var studentNames = [];
  var yearTag = "";
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastNamesValues = sheet.getRange(2, TRACK_LASTNAME_NDX, sheet.getMaxRows() - 1).getValues();
    var firstNamesValues = sheet.getRange(2, TRACK_FIRSTNAME_NDX, sheet.getMaxRows() - 1).getValues();

    // convert the array ignoring empty cells
    for(var j = 0; j < lastNamesValues.length; j++) {
      if(lastNamesValues[j][0] != "") {
        yearTag = getYearTag(i);
        if (!yearTag) continue;
        newName = lastNamesValues[j][0] + ", " + firstNamesValues[j][0] + " (" + yearTag + ")";
        if (studentNames.includes(newName)) {
          Logger.log("Duplicate: " + newName);
        }
        studentNames.push(newName);
      }
    }
  }
  studentNames = studentNames.sort();
  Logger.log(studentNames);
  return studentNames; 
}

/*
 * Given a name as formatted by the formatter, get an array with 
 * the student's [year_sheet_index, name_row_index]
 */
function findCellByName(name) {
  var sheets = SpreadsheetApp.openById(SHEET_TRACKER).getSheets();
  var nameArr = name.slice(0, -6).split(", ");
  var firstName = nameArr[1];
  var lastName = nameArr[0];

  var sheetNum = -1;
  var nameNum = -1;
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastNamesValues = sheet.getRange(1, TRACK_LASTNAME_NDX, sheet.getMaxRows() - 1).getValues();
    var firstNamesValues = sheet.getRange(1, TRACK_FIRSTNAME_NDX, sheet.getMaxRows() - 1).getValues();

    // convert the array ignoring empty cells
    for(var j = 0; j < lastNamesValues.length; j++) {
      //if(firstNamesValues[j][0][0] == firstName[0]) {Logger.log("%s, %s", firstNamesValues[j][0], firstName);}
      //if(lastNamesValues[j][0] == lastName) {Logger.log("%s, %s", lastNamesValues[j][0], lastName);}
      if(firstNamesValues[j][0] == firstName && lastNamesValues[j][0] == lastName) {
        sheetNum = i;
        nameNum = j;
        break;
      }
    }
    if (sheetNum != -1) break;
  }

  if (sheetNum == -1 && nameNum == -1) {
    Logger.log("Did not find name")
    Logger.log(name)
  }

  return([sheetNum, nameNum]);
}

// Update the name list for the Google Form 
function updateStudents() {
  var form = FormApp.getActiveForm();
  var namesList = form.getItemById(NAMES_ID).asListItem();

  // Generate name options from Match Tracker
  var studentNames = buildNameList();
  namesList.setChoiceValues(studentNames);
}

function getYearTag(sheetNum) {
  var tags = ["MS1", "MS2", "MS3", "MS4", "PA1", "PA2"]
  if (sheetNum > tags.length) return 0;
  return tags[sheetNum];
}

function GET_INFO(position, info) {
  var sheet = SpreadsheetApp.openById(SHEET_PEOPLE).getSheets()[0];
  name = ""
  email = ""
  switch(position) {
    case("CEO"):
      name = sheet.getRange(PEOPLE_CEO, 2).getValue();
      email = sheet.getRange(PEOPLE_CEO, 3).getValue();
      break;
    case("COO"):
      name = sheet.getRange(PEOPLE_COO, 2).getValue();
      email = sheet.getRange(PEOPLE_COO, 3).getValue();
      break;
    case("Webmaster"):
      name = sheet.getRange(PEOPLE_WEBMASTER, 2).getValue();
      email = sheet.getRange(PEOPLE_WEBMASTER, 3).getValue();
      break;
    case("GenPedManager"):
      name = sheet.getRange(PEOPLE_GEN_PED, 2).getValue();
      email = sheet.getRange(PEOPLE_GEN_PED, 3).getValue();
      break;
    case("WomenManager"):
      name = sheet.getRange(PEOPLE_WOMEN, 2).getValue();
      email = sheet.getRange(PEOPLE_WOMEN, 3).getValue();
      break;
    case("GeriDermManager"):
      name = sheet.getRange(PEOPLE_GERI_DERM, 2).getValue();
      email = sheet.getRange(PEOPLE_GERI_DERM, 3).getValue();
      break;
    case("DIMEManager"):
      name = sheet.getRange(PEOPLE_DIME, 2).getValue();
      email = sheet.getRange(PEOPLE_DIME, 3).getValue();
      break;
    case("ROCManager"):
      name = sheet.getRange(PEOPLE_ROC, 2).getValue();
      email = sheet.getRange(PEOPLE_ROC, 3).getValue();
      break;
    case("SMManager"):
      name = sheet.getRange(PEOPLE_SM, 2).getValue();
      email = sheet.getRange(PEOPLE_SM, 3).getValue();
      break;
    case("LayCouns"):
      name = sheet.getRange(PEOPLE_LAY, 2).getValue();
      email = sheet.getRange(PEOPLE_LAY, 3).getValue();
      break;
    case("ClassLists"):
      name = sheet.getRange(PEOPLE_CLASS, 2).getValue();
      email = sheet.getRange(PEOPLE_CLASS, 3).getValue();
      break;
    default:
      name = "Name Not Found";
      email = "Email Not Found";
      break;
  }
  switch(info) {
    case "Email": case "email":
      return email;
    case "Name": case "name":
      return name;
    default:
      return "Bad Lookup"
  }
}