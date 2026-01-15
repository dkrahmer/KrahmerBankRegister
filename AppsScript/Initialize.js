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

/*************************************************************************************
Version: 4.0.0
Release Date: 2022-03-31 (first release date: 2017-08-26)
Release notes:
  2026-01-14 (4.0.0)  - Add AI email bill handling
  2022-09-24 (3.0.0)  - Change to decending date transaction sort order
                        Add auto-archiving
                        Code reorg
                        Re-release as GPL
  2022-03-31 (2.0.0.0) - Automatically clear the "E" flag in status column when the amount is changed.
                         Use UTC dates for all date objects and change app to UTC zone
                         Fix sort after adding auto rows - needs flush before sorting since Google changes
                         Update next auto date in realtime as auto entries are created
                         More reliable method to update payee categories
  2021-11-29 - Change "R" column to "?" for status. Add estimate "E" flag in status column with automation support.
  2019-06-22 - Fix "drifting" day of month for recurring transactions due to Google DST bug
  2018-03-29 - Fix "drifting" day of month for recurring transactions due to GMT
  2017-09-23 - Added better support for auto rows when using mobile sheets
               onOpen() doesn't fire so use onEdit() as backup.
  2017-08-26 - Initial release
*************************************************************************************/

// ------------------ Begin Settings ------------------
const REGISTER_EMPTY_ROWS = 3;
const ESTIMATE_STATUS = "E"; // must match sheet data validation for the status columns

const REGISTER_SHEET_NAME = "Register";
const RECURRING_SHEET_NAME = "Recurring";
const ARCHIVE_SHEET_NAME = "Archive";
const PAYEES_SHEET_NAME = "Payees";

const WEB_SERVICE_PASSPHRASE = PropertiesService.getScriptProperties().getProperty('WEB_SERVICE_PASSPHRASE') ?? '';
// ------------------ End Settings ------------------

const WAIT_LOCK_TIMEOUT = 10000;
const REGISTER_FOOTER_ROW_COUNT = 2;
const ARCHIVE_FOOTER_ROW_COUNT = 1;

const SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

let REGISTER_COL_COUNT = 0;
const REGISTER_COL_DATE = ++REGISTER_COL_COUNT;
const REGISTER_COL_NUM = ++REGISTER_COL_COUNT;
const REGISTER_COL_PAYEE = ++REGISTER_COL_COUNT;
const REGISTER_COL_CATEGORY = ++REGISTER_COL_COUNT;
const REGISTER_COL_WITHDRAWAL = ++REGISTER_COL_COUNT;
const REGISTER_COL_DEPOSIT = ++REGISTER_COL_COUNT;
const REGISTER_COL_STATUS = ++REGISTER_COL_COUNT;
const REGISTER_COL_BALANCE = ++REGISTER_COL_COUNT;
const REGISTER_COL_SORT = ++REGISTER_COL_COUNT;
const REGISTER_COL_NOTES = ++REGISTER_COL_COUNT;

let RECURRING_COL_COUNT = 0;
const RECURRING_COL_ENABLED = ++RECURRING_COL_COUNT;
const RECURRING_COL_FREQUENCY = ++RECURRING_COL_COUNT;
const RECURRING_COL_NEXT_DATE = ++RECURRING_COL_COUNT;
const RECURRING_COL_END_DATE = ++RECURRING_COL_COUNT;
const RECURRING_COL_DAYS_AHEAD = ++RECURRING_COL_COUNT;
const RECURRING_COL_PAYEE = ++RECURRING_COL_COUNT;
const RECURRING_COL_CATEGORY = ++RECURRING_COL_COUNT;
const RECURRING_COL_DEBIT = ++RECURRING_COL_COUNT;
const RECURRING_COL_CREDIT = ++RECURRING_COL_COUNT;
const RECURRING_COL_NOTES = ++RECURRING_COL_COUNT;
const RECURRING_COL_STATUS = ++RECURRING_COL_COUNT;
const RECURRING_COL_NOTES_PRIVATE = ++RECURRING_COL_COUNT;

let PAYEES_COL_COUNT = 0;
const PAYEES_COL_PAYEE = ++PAYEES_COL_COUNT;
const PAYEES_COL_CATEGORY = ++PAYEES_COL_COUNT;
const PAYEES_COL_EMAIL_AI_RULES = ++PAYEES_COL_COUNT;

function initialize() {
  console.log(`${getFuncName()}...`);
  saveSheetKey();
  deleteRegister();
  deleteArchive();
  deleteRecurring();
  deletePayees();
  createTriggers();
}

function saveSheetKey() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  SCRIPT_PROP.setProperty("key", sheet.getId());
}

function createTriggers() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  let scriptName;

  const triggerFunctionNames = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());

  scriptName = "onEditCustom";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .forSpreadsheet(sheet).onEdit()
      .create();
    console.log(`Scheduled ${scriptName}`);
  }

  scriptName = "processRecurring";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .timeBased().everyDays(1).atHour(1).nearMinute(0)
      .create();
    console.log(`Scheduled ${scriptName}`);
  }

  scriptName = "autoArchive";
  if (!triggerFunctionNames.includes(scriptName)) {
    ScriptApp.newTrigger(scriptName)
      .timeBased().everyDays(1).atHour(2).nearMinute(0)
      .create();
    console.log(`Scheduled ${scriptName}`);
  }
}

function deleteRegister() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const registerSheet = sheet.getSheetByName(REGISTER_SHEET_NAME);

  const headerRows = registerSheet.getFrozenRows();
  registerSheet.deleteRows(headerRows + 1, registerSheet.getMaxRows() - headerRows - REGISTER_FOOTER_ROW_COUNT);
  validateRegisterSheet();
}

function deleteArchive() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const archiveSheet = sheet.getSheetByName(ARCHIVE_SHEET_NAME);

  const headerRows = archiveSheet.getFrozenRows();
  archiveSheet.deleteRows(headerRows + 1, archiveSheet.getMaxRows() - headerRows - ARCHIVE_FOOTER_ROW_COUNT);
}

function deleteRecurring() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const recurringSheet = sheet.getSheetByName(RECURRING_SHEET_NAME);

  const headerRows = recurringSheet.getFrozenRows();
  const range = recurringSheet.getRange(headerRows + 1, 1, recurringSheet.getMaxRows() - headerRows, recurringSheet.getMaxColumns());
  range.clearContent();
}

function deletePayees() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  const payeesSheet = sheet.getSheetByName(PAYEES_SHEET_NAME);

  const headerRows = payeesSheet.getFrozenRows();
  const range = payeesSheet.getRange(headerRows + 1, 1, payeesSheet.getMaxRows() - headerRows, payeesSheet.getMaxColumns());
  range.clearContent();
}
