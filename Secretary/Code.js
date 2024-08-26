const SHEETS_ID = {
    TRACKER: "10e68w1DkTm4kdXJIcMUeIUH5_KFP1uUgKv5SB5RHXDU",  // Match tracker
    PEOPLE: "1R7sskPPhNi6Mhitz1-FHESdJhaJuKHM_o8oUJHSp9EQ"    // Contact info
  };

function processAttendanceAndNotifyBoardMembers() {
  const sheet = SpreadsheetApp.openById(SHEETS_ID.TRACKER).getSheets()[0];
  const dateAllColumn = sheet.getRange(1, TRACK_INDEX.DATE_ALL, sheet.getLastRow(), 1).getValues().flat().filter(String);
  
  const boardMembers = [
    { name: 'John Doe', email: 'john@example.com' },
    { name: 'Jane Smith', email: 'jane@example.com' },
    // Add more board members here
  ];

  const monthCounts = {};
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

  dateAllColumn.forEach(dateString => {
    const dates = dateString.split(',').map(d => d.trim());
    dates.forEach(date => {
      const month = new Date(date).getMonth();
      monthCounts[month] = (monthCounts[month] || 0) + 1;
    });
  });

  months.forEach((month, index) => {
    if (!monthCounts[index]) {
      boardMembers.forEach(member => {
        sendNoAttendanceEmail(member, month);
      });
    }
  });
}

function sendNoAttendanceEmail(member, month) {
  const subject = `No Clinic Attendance in ${month}`;
  const body = `Dear ${member.name},\n\nThis is to inform you that you did not attend any clinics in ${month}. Please ensure you maintain regular attendance.\n\nBest regards,\nSecretary`;
  
  MailApp.sendEmail(member.email, subject, body);
}
