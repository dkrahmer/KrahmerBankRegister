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

function handleRecurringSheetEdit(e) {
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
    // Update payee categories
    if (e && [RECURRING_COL_PAYEE, RECURRING_COL_CATEGORY].includes(e.range.getColumn())) {
      const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
      const recurringSheet = sheet.getSheetByName(RECURRING_SHEET_NAME);

      const changedRowRange = recurringSheet.getRange(e.range.getRow(), RECURRING_COL_PAYEE, 1, 2); // get the 2 columns for payee & category
      const changedRow = changedRowRange.getValues()[0];
      const newPayee = changedRow[0];
      const newCategory = changedRow[1];
      const payeesSheet = sheet.getSheetByName(PAYEES_SHEET_NAME);
      const payeesRows = payeesSheet.getDataRange().getValues();
      let foundPayee = false;

      if (!!newPayee && e.range.getColumn() === RECURRING_COL_PAYEE) {
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
            // Set the category to the recurring sheet
            recurringSheet.getRange(e.range.getRow(), RECURRING_COL_CATEGORY).setValue(category);
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
        SpreadsheetApp.flush();
      }
    }

    processRecurring();
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}

function processRecurring() {
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
    const recurringSheet = sheet.getSheetByName(RECURRING_SHEET_NAME);
    const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);
    const today = new Date();
    const rowCount = recurringSheet.getMaxRows();
    let sortOnComplete = false;

    do {
      sortOnComplete = false;
      for (let rowNum = 2; rowNum <= rowCount; rowNum++) {
        const rowValues = recurringSheet.getRange(rowNum, 1, 1, RECURRING_COL_COUNT).getValues()[0];
        const values = {
          enabled: rowValues[RECURRING_COL_ENABLED - 1],
          frequency: rowValues[RECURRING_COL_FREQUENCY - 1],
          nextDateUtc: rowValues[RECURRING_COL_NEXT_DATE - 1],
          endDateUtc: rowValues[RECURRING_COL_END_DATE - 1],
          daysAhead: rowValues[RECURRING_COL_DAYS_AHEAD - 1],
          payee: rowValues[RECURRING_COL_PAYEE - 1],
          category: rowValues[RECURRING_COL_CATEGORY - 1],
          debit: rowValues[RECURRING_COL_DEBIT - 1],
          credit: rowValues[RECURRING_COL_CREDIT - 1],
          notes: rowValues[RECURRING_COL_NOTES - 1],
          status: rowValues[RECURRING_COL_STATUS - 1],
          notesPrivate: rowValues[RECURRING_COL_NOTES_PRIVATE - 1]
        }

        if (!values.enabled || !values.frequency)
          continue;

        if (!(values.nextDateUtc instanceof Date) || isNaN(values.nextDateUtc))
          continue;

        if ((values.endDateUtc instanceof Date) && today > values.endDateUtc)
          continue;

        values.nextDateUtc = new Date(values.nextDateUtc.toISOString().substring(0, 10) + "Z");

        const updateRange = recurringSheet.getRange(rowNum, RECURRING_COL_NEXT_DATE);
        const updateNextDate = (newNextDateUtc) => { updateRange.setValue(newNextDateUtc); }
        const newNextDateUtc = insertRecurringTransactions(registerSheet, values, updateNextDate);

        if (newNextDateUtc != null && newNextDateUtc != values.nextDateUtc)
          sortOnComplete = true;
      }
      if (sortOnComplete)
        sortRegisterSheet();
    } while (sortOnComplete > 0); // check in case recurring was changed by the user while processing
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}

function insertRecurringTransactions(targetSheet, values, updateNextDate) {
  console.log(`${getFuncName()} - ${values?.payee}...`);
  const nowLocal = new Date();
  const endDateUtc = new Date(Date.UTC(nowLocal.getFullYear(), nowLocal.getMonth(), nowLocal.getDate() + values.daysAhead)); // cast local date as UTC
  let nextDateUtc = values.nextDateUtc;
  while (nextDateUtc instanceof Date && !isNaN(nextDateUtc) && nextDateUtc <= endDateUtc) {
    // Use addRegisterEntry to add a transaction
    addRegisterEntry({
      date: nextDateUtc,
      payee: values.payee,
      category: values.category,
      debit: values.debit,
      credit: values.credit,
      status: values.status,
      notes: values.notes,
      source: "auto"
    }, false);
    nextDateUtc = getNextDateUtc(nextDateUtc, values.frequency);
    if (updateNextDate)
      updateNextDate(nextDateUtc);
  }
  return nextDateUtc;
}
