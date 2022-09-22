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

function onEditCustom(e) {
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
    const sourceSheetName = e.source.getSheetName();

    if (sourceSheetName == REGISTER_SHEET_NAME) {
      handleRegisterSheetEdit(e);
    }
    else if (sourceSheetName == RECURRING_SHEET_NAME) {
      handleRecurringSheetEdit(e);
    }
  }
  finally {
      SpreadsheetApp.flush();
      lock.releaseLock();
  }
}

function onOpen() {
  console.log(`${getFuncName()}...`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  var entries = [
    {name:"Archive Now", functionName:"autoArchive"},
    {name:"Process Recurring", functionName:"processRecurring"},
    {name:"Sort Register", functionName:"sortRegisterSheet"},
    {name:"Sort Archive", functionName:"sortArchiveSheet"}
  ];
  sheet.addMenu("Account Register Actions", entries);
}
