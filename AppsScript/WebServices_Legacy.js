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

// This file contains legacy code that is no longer used but may be useful to add web service functionality in the future

/*
function doGet(e) {
  const functionName = e.parameter["function"] ?? "";
  const passphrase = e.parameter["passphrase"] ?? "";

  if (passphrase != WEB_SERVICE_PASSPHRASE.replace(/[^a-zA-Z0-9]/g, '')) {
    return ContentService
      .createTextOutput(JSON.stringify({"success": false, "reason": "invalid passphrase"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const sheet = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));

    if (functionName == "abc") {
      abc();

      return ContentService
        .createTextOutput(JSON.stringify({"success": true, "reason": "done"}))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({"success": false, "reason": `error: "Invalid Function"`}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({"success": false, "reason": `error: ${JSON.stringify(e)}`}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
*/

// This code allows entries to be inserted from other applications by running as a web service.
// Here is a sample JavaScript to utilize this web service once it is deployed.
// Follow step by step instructions below to deploy.
//
// You must be logged into a Google account that has write access to the account register service
// and for the the following JavaScript code to work.
// Include the following commented code in your own JavaScript app.
/*
// JavaScript code to be used in your custom web app. Works with TamperMonkey.
insertToCheckRegister({ "Date": paymentDate, "Payee/Transaction Description": payeeName, "Withdrawal, Payment (-)": paymentAmount, "Notes": notes});

function insertToBankRegister(data) {
    const macroUrl = "https://script.google.com/macros/s/<Put your macro ID here>/dev";
    const iframeUrl = macroUrl + "?" + $.param(data);
    const iframeSortUrl = macroUrl + "?sort";

    const iframe = document.createElement('iframe');
    iframe.style.display = "none";
    document.body.appendChild(iframe);
    const iframeDoc = iframe.contentWindow.document;

    const iframe2 = document.createElement('iframe');
    iframe2.style.display = "none";
    document.body.appendChild(iframe2);
    const iframeDoc2 = iframe2.contentWindow.document;

    iframeDoc.open();
    iframeDoc.write('<script>window.location = "' + iframeUrl + '"</script>');
    iframeDoc.close();

    iframe.onload = function() {
        iframeDoc2.open();
        iframeDoc2.write('<script>window.location = "' + iframeSortUrl + '"</script>');
        iframeDoc2.close();
    };
}

function openBankRegister () {
    setTimeout(function() { window.open("https://docs.google.com/spreadsheets/d/<Put your Google Sheet ID here>/edit"); }, 0);
}
*/

/*
//  2. Run initialize()
//
//  3. Publish > Deploy as web app
//    - enter Project Version name and click 'Save New Version'
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously)
//
//  4. Copy the 'Current web app URL' and post this in your form/script action
//
//  5. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  if (e.queryString == "sort") {
    const doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    const sheet = doc.getSheetByName(REGISTER_SHEET_NAME);
    sheet.sort(9, true);
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "sorted": true}))
          .setMimeType(ContentService.MimeType.JSON);
  }
  else {
    return handleResponse(e);
  }
}

function doPost(e){
  return handleResponse(e);
}

function handleResponse(e) {
  const lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.

  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    const doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    const sheet = doc.getSheetByName(REGISTER_SHEET_NAME);

    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    const headRow = 2;
    const headers = sheet.getRange(headRow, 1, headRow, sheet.getLastColumn()).getValues()[0];
    //const nextRow = sheet.getLastRow()+1; // get next row
    const nextRow = getLastPopulatedRow(sheet, 1) + 2;
    const row = [];
    // loop through the header columns
    const headerCount = 0;
    for (i in headers) {
      headerCount++;
      const headerName = headers[i];
      if (headerName == undefined || headerName== "")
        break;
      if (headerName == "Timestamp"){ // special case if you include a 'Timestamp' column
        row.push(new Date());
      }
      else { // else use header name to get data
        const cellData = e.parameter[headerName];
        if (cellData != undefined && cellData != "") {
          sheet.getRange(nextRow, headerCount, 1, 1).setValue(cellData);
        }
      }
    }

    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}
*/