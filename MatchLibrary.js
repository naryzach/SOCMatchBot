/**
 * Match Library for Student Outreach Clinics
 * 
 * This script contains utility functions for managing student matches and information
 * for various outreach clinic programs. It provides functionality for building name lists,
 * finding student information, updating forms, and retrieving staff information.
 * 
 * The library works with multiple sheets:
 * - Match tracker sheet: Contains student sign-up and match information
 * - People sheet: Contains staff and management contact information
 * 
 * Important:
 * - Ensure all sheet IDs and indices are up to date
 * - This library is used by multiple clinic scripts, so changes may affect multiple programs
 * - Verify that all referenced spreadsheet ranges and column indices are correct
 */

// Match tracker sheet
const TRACK_INDEX = {
    LASTNAME: 1,
    FIRSTNAME: 2,
    SIGNUPS: 3,
    MATCHES: 4,
    NOSHOW: 5,
    CXLLATE: 6,
    CXLEARLY: 7,
    DATE: 8,
    DATE_ALL: 9
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
    CLASS: 12,
    SECRETARY: 13
  };

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
  
    for (let sheetIndex = 0; sheetIndex < 6; sheetIndex++) { // 6 classes (4 MS, 2 PA)
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
    Logger.log(studentNames);
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
    const positions = {
      CEO: PEOPLE_INDEX.CEO,
      COO: PEOPLE_INDEX.COO,
      Webmaster: PEOPLE_INDEX.WEBMASTER,
      GenPedManager: PEOPLE_INDEX.GEN_PED,
      WomenManager: PEOPLE_INDEX.WOMEN,
      GeriDermManager: PEOPLE_INDEX.GERI_DERM,
      DIMEManager: PEOPLE_INDEX.DIME,
      ROCManager: PEOPLE_INDEX.ROC,
      SMManager: PEOPLE_INDEX.SM,
      LayCouns: PEOPLE_INDEX.LAY,
      ClassLists: PEOPLE_INDEX.CLASS,
      Secretary: PEOPLE_INDEX.SECRETARY
    };
  
    const rowIndex = positions[position];
    if (!rowIndex) {
      return info === "email" ? "Email Not Found" : "Name Not Found";
    }
  
    const name = sheet.getRange(rowIndex, 2).getValue();
    const email = sheet.getRange(rowIndex, 3).getValue();
  
    return info.toLowerCase() === "email" ? email : name;
  }