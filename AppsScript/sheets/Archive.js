/************************************************************************************************************
Krahmer Bank Register
Copyright 2022 Douglas Krahmer

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

function sortArchiveSheet() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const archiveSheet = sheet.getSheetByName(ARCHIVE_SHEET_NAME);

  SpreadsheetApp.flush();

  // Sort the register
  archiveSheet.sort(REGISTER_COL_SORT, false);
}

function autoArchive() {
  console.log("___...");
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(WAIT_LOCK_TIMEOUT);
  }
  catch (ex) {
    console.error(ex);
    return;
  }

  try {
    const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);
    const registerHeaderRows = registerSheet.getFrozenRows();
    const archiveSheet = sheet.getSheetByName(ARCHIVE_SHEET_NAME);
    const archiveAfterDays = archiveSheet.getRange("E1").getValue() ?? 365;
    const now = new Date();
    const maxArchiveDate = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - archiveAfterDays));
    const skipDatesBefore = new Date(1901, 1, 1);

    for (let rowNum = registerSheet.getMaxRows(); rowNum > registerHeaderRows; rowNum--) {
      const rowDate = registerSheet.getRange(rowNum, REGISTER_COL_DATE).getValue();

      if (!(rowDate instanceof Date) || isNaN(rowDate))
        continue;
        
      if (rowDate < skipDatesBefore)
        continue; // this is a template row that should never be deleted
        
      if (rowDate > maxArchiveDate)
        break; // all done
      
      archiveRow(registerSheet, archiveSheet, rowNum);
    }
    
    validateRegisterSheet();
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}

function archiveRow(registerSheet, archiveSheet, rowNum) {
  console.log(`${getFuncName()}...`);
  const archiveHeaderRows = archiveSheet.getFrozenRows();

  const registerRow = registerSheet.getRange(rowNum, 1, 1, REGISTER_COL_COUNT);
  archiveSheet.insertRowBefore(archiveHeaderRows + 1);
  const archiveRow = archiveSheet.getRange(archiveHeaderRows + 1, 1, 1, REGISTER_COL_COUNT);
  registerRow.copyTo(archiveRow, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
  registerSheet.deleteRow(rowNum);
}
