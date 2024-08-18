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