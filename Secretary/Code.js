/**
 * This script is designed to process attendance data and notify board members of any missing attendance in a given month.
 * It reads data from a spreadsheet, calculates the attendance for each month, and sends emails to board members who have no attendance records for that month.
 * 
 * @author [Ryan Gustafson]
 * @version 1.0
 * @since 2024-07-23
 */

const SHEETS_ID = {
    TRACKER: "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU",  // Match tracker
    PEOPLE: "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"    // Contact info
};

/**
 * Creates installable triggers for the form submit and daily update functions.
 * This function should be run manually after updating the Match Tracker.
 */
function createTriggers() {
  // Check if triggers are already set 
  var currentTriggers = ScriptApp.getProjectTriggers();
  if (currentTriggers.length > 0) {
    Logger.log("Triggers already set.");
    return;
  }

  ScriptApp.newTrigger("checkAttendance")
    .timeBased()
    .atHour(9)  // Run at 1 AM to ensure it's the first of the month in all timezones
    .onMonthDay(1)
    .everyMonth()
    .create();
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
 * Checks the attendance data and sends emails to board members who have no attendance records for a given month.
 */
function checkAttendance() {
  const tracker = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets();
  const lastMonth = new Date();
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const lastMonthNumber = lastMonth.getMonth();
  const lastMonthName = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][lastMonthNumber];

  const boardMembers = [
    //{ name: GET_INFO("CEO", "name")), email: GET_INFO("CEO", "email") },
    //{ name: GET_INFO("COO", "name")), email: GET_INFO("COO", "email") },
    //{ name: GET_INFO("Chairman", "name")), email: GET_INFO("Chairman", "email") },
    { name: GET_INFO("Webmaster", "name"), email: GET_INFO("Webmaster", "email") },
    //{ name: GET_INFO("Treasurer", "name")), email: GET_INFO("Treasurer", "email") },
    //{ name: GET_INFO("Secretary", "name")), email: GET_INFO("Secretary", "email") }
  ];

  boardMembers.forEach(member => {
    const [memberSheet, memberRow] = findCellByName(adjustName(member.name));
    const trackSheet = tracker[memberSheet];

    if (memberRow !== -1) {
      const memberDates = trackSheet.getRange(memberRow, TRACK_INDEX.DATE_ALL).getValue().split(',').map(d => d.trim()) || "";
      if (memberDates === "" || !memberDates.some(date => new Date(date).getMonth() === lastMonthNumber)) {
        const subject = `No Clinic Attendance in ${lastMonthName}`;
        const htmlTemplate = HtmlService.createTemplateFromFile('attendanceEmail');
        htmlTemplate.name = member.name.slice(0, -5);
        htmlTemplate.month = lastMonthName;
        htmlTemplate.feedback_email = GET_INFO("Webmaster", "email");
        const htmlBody = htmlTemplate.evaluate().getContent();
        MailApp.sendEmail({
          to: member.email,
          subject: subject,
          htmlBody: htmlBody,
          name: "Secretary Assistant",
          replyTo: GET_INFO("Secretary", "email")
        });
      }
    }
  });
}

function adjustName(name) {
  const nameParts = name.slice(0, -5).split(' ');
  const lastName = nameParts.pop();
  const firstName = nameParts.join(' ');
  return `${lastName}, ${firstName} (MSX)`;
}