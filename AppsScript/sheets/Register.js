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

function updatePendingRegisterEntry(registerRowData) {
  console.log(`${getFuncName()}...`);
  let { maxDate, payee, amount, mode, dueDate, status, appendNote, informationOnlyEntry } = registerRowData;
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);

  SpreadsheetApp.flush();

  const rows = registerSheet.getDataRange().getValues();
  let threshold = maxDate ? new Date(maxDate) : new Date();
  threshold.setHours(0, 0, 0, 0);

  let targetDueDate = null;
  const twoWeeksMs = 14 * 24 * 60 * 60 * 1000;
  if (dueDate) {
    targetDueDate = new Date(dueDate);
    targetDueDate.setHours(0, 0, 0, 0);
  }

  let candidateIndex = -1;
  let minDate = null;

  for (let i = 1; i < rows.length; i++) { // Skip header row (assuming row 0 is header)
    const row = rows[i];
    const dateValue = row[REGISTER_COL_DATE - 1];
    if (!dateValue || !(dateValue instanceof Date))
      continue; // Skip rows with empty date

    const rowDate = new Date(dateValue);
    rowDate.setHours(0, 0, 0, 0);

    if (rowDate <= threshold)
      break; // Stop searching if date is threshold or older

    // Match payee (case-insensitive, trimmed)
    const rowPayee = (row[REGISTER_COL_PAYEE - 1] || '').toString().trim().toLowerCase();
    const targetPayee = payee.trim().toLowerCase();
    if (rowPayee !== targetPayee)
      continue;

    // If dueDate provided, check within 2 weeks
    if (targetDueDate && Math.abs(rowDate.getTime() - targetDueDate.getTime()) > twoWeeksMs)
      continue;

    // Track the one with earliest date
    if (minDate === null || rowDate < minDate) {
      minDate = rowDate;
      candidateIndex = i;
    }
  }

  if (candidateIndex === -1) {
    throw new Error(`No matching pending entry found for payee: ${payee}`);
  }

  // Found the matching row - update it
  const row = rows[candidateIndex];

  // Update amount if provided
  if (amount !== undefined) {
    let withdrawal = parseFloat(row[REGISTER_COL_WITHDRAWAL - 1]) ?? 0;

    mode = mode ?? 'replace'; // Default to 'replace' if not provided

    if (mode === 'replace') {
      withdrawal = amount;
    }
    else if (mode === 'add') {
      withdrawal = `=${withdrawal}+(${amount})`;
    }
    else {
      throw new Error('Invalid mode: must be "replace" or "add"');
    }

    // Update the register with the withdrawal
	registerSheet.getRange(candidateIndex + 1, REGISTER_COL_WITHDRAWAL, 1, 1).setValue(withdrawal);

    // If informationOnlyEntry flag is set, also set deposit to the same amount
    if (informationOnlyEntry) {
      registerSheet.getRange(candidateIndex + 1, REGISTER_COL_DEPOSIT, 1, 1).setValue(amount);
    }
  }

  // Set status if provided
  if (status !== undefined) {
    row[REGISTER_COL_STATUS - 1] = status;
    registerSheet.getRange(candidateIndex + 1, REGISTER_COL_STATUS, 1, 1).setValue(status);
  }

  // Append note if provided
  if (appendNote) {
    let notes = row[REGISTER_COL_NOTES - 1] ?? '';

    // Check if the appendNote already exists in notes
    const notePattern = new RegExp(`(^|\\s-\\s)${appendNote.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(\\s\\((\\d+)\\))?(\\s-\\s|$)`);
    const match = notes.match(notePattern);

    if (match) {
      // Note exists, increment or add count
      const currentCount = match[3] ? parseInt(match[3]) : 1;
      const newCount = currentCount + 1;
      const replacement = match[1] + appendNote + ` (${newCount})` + match[4];
      notes = notes.replace(notePattern, replacement);
    }
    else {
      // Note doesn't exist, append normally
      if (notes)
        notes += ' - ';
      notes += appendNote;
    }

    row[REGISTER_COL_NOTES - 1] = notes;
    registerSheet.getRange(candidateIndex + 1, REGISTER_COL_NOTES, 1, 1).setValue(notes);
  }

  // Overwrite date if provided
  if (targetDueDate) {
    registerSheet.getRange(candidateIndex + 1, REGISTER_COL_DATE, 1, 1).setValue(targetDueDate);
  }

  sortRegisterSheet();

  SpreadsheetApp.flush();
  console.log(`${getFuncName()} complete.`);
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
      if (changedCell.getValue() === ESTIMATE_STATUS)
        changedCell.setValue("");
    }

    // Update payee categories
    if (e && [REGISTER_COL_PAYEE, REGISTER_COL_CATEGORY].includes(e.range.getColumn())) {
      const changedRowRange = registerSheet.getRange(e.range.getRow(), REGISTER_COL_PAYEE, 1, 2); // get the 2 columns for payee & category
      const changedRow = changedRowRange.getValues()[0];
      const newPayee = changedRow[0];
      const newCategory = changedRow[1];
      const payeesSheet = sheet.getSheetByName(PAYEES_SHEET_NAME);
      const payeesRows = payeesSheet.getDataRange().getValues();
      let foundPayee = false;

      if (!!newPayee && e.range.getColumn() === REGISTER_COL_PAYEE) {
        // the payee was set and the category is empty. Try to fill in the category...
        for (
          let updateRowNumber = 1;
          updateRowNumber < payeesRows.length;
          updateRowNumber++
        ) {
          const payeesRow = payeesRows[updateRowNumber];
          let payee = payeesRow[0];
          if (!payee)
            break;

          if (payee === newPayee) {
            foundPayee = true;
            let category = payeesRow[1];
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
          updateRowNumber < payeesRows.length;
          updateRowNumber++
        ) {
          const payeesRow = payeesRows[updateRowNumber];
          payee = payeesRow[0];
          category = payeesRow[1];

          foundPayee = payee == newPayee;
          if (foundPayee || !payee || !category)
            break;
        }

        if (category != newCategory)
          payeesSheet
            .getRange(updateRowNumber + 1, 2)
            .setValue(newCategory);

        if (!foundPayee) {
          payeesSheet
            .getRange(updateRowNumber + 1, 1)
            .setValue(newPayee);
          // Sort the payees after adding a new payee
          payeesSheet.sort(1, true);
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
