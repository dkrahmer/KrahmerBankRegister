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

function getPayees() {
  const payeesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYEES_SHEET_NAME);
  const payeesRows = payeesSheet.getDataRange().getValues();
  
  // Skip header row (assuming row 0 is headers)
  const payees = [];
  for (let i = 1; i < payeesRows.length; i++) {
    const row = payeesRows[i];
    if (row[PAYEES_COL_PAYEE - 1]) {  // Only add if payee is present
      payees.push({
        payee: row[PAYEES_COL_PAYEE - 1],
        category: row[PAYEES_COL_CATEGORY - 1],
        emailAiRules: row[PAYEES_COL_EMAIL_AI_RULES - 1]
      });
    }
  }
  
  return payees;
}