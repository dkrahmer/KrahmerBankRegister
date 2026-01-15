/************************************************************************************************************
Krahmer Bank Register
Copyright 2026 Douglas Krahmer

This file is part of Krahmer Bank Register.

Krahmer Bank Register is free software: you can redistribute it and/or modify it under the terms of the
GNU General Public License as published by the Free Software Foundation, either version 3 of the License,
or (at your option) any later version.

Krahmer Bank Register is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with Krahmer Bank Register.
If not, see <https://www.gnu.org/licenses/>.
************************************************************************************************************/

/**
 * Write a log entry to the Log sheet
 * @param {string} level - Log level (INFO, ERROR, DEBUG)
 * @param {string} message - Log message
 * @param {Object} data - Optional data object to log
 */
function writeLog(level, message, data = null) {
  try {
    const ss = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    const logSheet = ss.getSheetByName("Log");

    if (!logSheet) {
      // If Log sheet doesn't exist, silently fail
      return;
    }

    const timestamp = new Date();
    const dataStr = data ? JSON.stringify(data) : "";

    logSheet.appendRow([timestamp, level, message, dataStr]);

    // Keep only last 1000 entries
    const maxRows = 1000;
    if (logSheet.getLastRow() > maxRows + 1) { // +1 for header
      logSheet.deleteRows(2, logSheet.getLastRow() - maxRows - 1);
    }
  }
  catch (err) {
    // Silently fail if logging fails - don't break the main function
  }
}
