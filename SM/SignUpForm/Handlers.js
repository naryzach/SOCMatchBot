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
  function handlePreliminaryMatch(form, dateString, clinicDate, spreadsheet) {
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
 * Creates a PDF of the match list for a given date.
 * @param {Date} date - The date of the clinic.
 * @return {GoogleAppsScript.Drive.File} The created PDF file.
 */
function makeMatchPDF(date) {
  // Format the PDF name
  const pdfName = `MatchList_${date.toISOString().split('T')[0]}.pdf`;
  
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