/************************************************************************************************************
Krahmer Account Register
Copyright 2022 Douglas Krahmer

This file is part of Krahmer Account Register.

Krahmer Account Register is free software: you can redistribute it and/or modify it under the terms of the 
GNU General Public License as published by the Free Software Foundation, either version 3 of the License, 
or (at your option) any later version.

Krahmer Account Register is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with Krahmer Account Register.
If not, see <https://www.gnu.org/licenses/>.
************************************************************************************************************/

function sortRegisterSheet() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);

  SpreadsheetApp.flush();

  // Sort the register
  registerSheet.sort(REGISTER_COL_SORT, false);
}

function validateRegisterSheet() {
  console.log(`${getFuncName()}...`);
  sortRegisterSheet();
  
  SpreadsheetApp.flush();
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);
  const headerRows = registerSheet.getFrozenRows();
  const startingRowCount = registerSheet.getMaxRows() - headerRows;

  // ensure minimum empty rows at the top
  let rowsToAdd = REGISTER_EMPTY_ROWS;
  if (startingRowCount < REGISTER_EMPTY_ROWS)
    addRegisterRows(REGISTER_EMPTY_ROWS - startingRowCount);

  const emptyRowRange = registerSheet.getRange(headerRows + 1, REGISTER_COL_SORT, REGISTER_EMPTY_ROWS, 1);
  const rowValues = emptyRowRange.getValues();
  for (let i = 0; i < rowValues.length; i++) {
    if (rowValues[i][0] || rowsToAdd <= 0)
      break;

    rowsToAdd--;
  }
  
  addRegisterRows(rowsToAdd, headerRows + 1);

  if (rowsToAdd > 0)
    sortRegisterSheet();
}

function addRegisterRows(rowsToAdd, rowToAddBefore) {
  console.log(`${getFuncName()}...`);
  if (rowsToAdd <= 0)
    return;

  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);
  rowToAddBefore = rowToAddBefore ?? (registerSheet.getFrozenRows() + 1);
  registerSheet.insertRowsBefore(rowToAddBefore, rowsToAdd);

  if (rowToAddBefore > registerSheet.getMaxRows())
    rowToAddBefore = registerSheet.getMaxRows();

  // copy formulas to the new rows
  SpreadsheetApp.flush();
  const templateRange = registerSheet.getRange(rowToAddBefore + rowsToAdd, REGISTER_COL_BALANCE, 1, REGISTER_COL_SORT - REGISTER_COL_BALANCE + 1);
  const targetRange = registerSheet.getRange(rowToAddBefore, REGISTER_COL_BALANCE, rowsToAdd, REGISTER_COL_SORT - REGISTER_COL_BALANCE + 1);
  templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
}

function handleRegisterSheetEdit(e) {
  console.log(`${getFuncName()}...`);
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

    // Clear the ESTIMATE_STATUS flag if this is an updated amount edit
    if (e
        && (!!e.oldValue || !!e.value)
        && [REGISTER_COL_WITHDRAWAL, REGISTER_COL_DEPOSIT].includes(e.range.getColumn())) {
      const changedCell = registerSheet.getRange(e.range.getRow(), REGISTER_COL_STATUS);
      if (changedCell.getValue() == ESTIMATE_STATUS)
        changedCell.setValue("");
    }

    // Update payee categories
    if (e && [REGISTER_COL_PAYEE, REGISTER_COL_CATEGORY].includes(e.range.getColumn())) {
      const changedRowRange = registerSheet.getRange(e.range.getRow(), REGISTER_COL_PAYEE, 1, 2); // get the 2 columns for payee & category
      const changedRow = changedRowRange.getValues()[0];
      const newPayee = changedRow[0];
      const newCategory = changedRow[1];
      const payeeCategoriesSheet = sheet.getSheetByName(PAYEE_CATEGORIES_SHEET_NAME);
      const payeeCategoriesRows = payeeCategoriesSheet.getDataRange().getValues();
      let foundPayee = false;

      if (!!newPayee && e.range.getColumn() == REGISTER_COL_PAYEE) {
        // the payee was set and the category is empty. Try to fill in the category...
        for (
          updateRowNumber = 1;
          updateRowNumber < payeeCategoriesRows.length;
          updateRowNumber++
        ) {
          const payeeCategoriesRow = payeeCategoriesRows[updateRowNumber];
          payee = payeeCategoriesRow[0];
          if (!payee) break;

          if (payee == newPayee) {
            foundPayee = true;
            category = payeeCategoriesRow[1];
            // Set the category to the register
            registerSheet.getRange(e.range.getRow(), REGISTER_COL_CATEGORY).setValue(category);
            break;
          }
        }
      }

      if (!foundPayee && !!newPayee && !!newCategory) {
        let payee = newPayee;
        let category = newCategory;
        let updateRowNumber;

        for (
          updateRowNumber = 1;
          updateRowNumber < payeeCategoriesRows.length;
          updateRowNumber++
        ) {
          const payeeCategoriesRow = payeeCategoriesRows[updateRowNumber];
          payee = payeeCategoriesRow[0];
          category = payeeCategoriesRow[1];

          foundPayee = payee == newPayee;
          if (foundPayee || !payee || !category) break;
        }

        if (category != newCategory)
          payeeCategoriesSheet
            .getRange(updateRowNumber + 1, 2)
            .setValue(newCategory);

        if (!foundPayee) {
          payeeCategoriesSheet
            .getRange(updateRowNumber + 1, 1)
            .setValue(newPayee);
          // Sort the payees after adding a new payee
          payeeCategoriesSheet.sort(1, true);
        }
      }
    }

    validateRegisterSheet();
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}
