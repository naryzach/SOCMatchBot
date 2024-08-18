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