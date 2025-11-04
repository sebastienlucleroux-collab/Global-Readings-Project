function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getScriptUrl() {
  var url = ScriptApp.getService().getUrl();
  return url;
}

function doGet(e) {
  if (!e.parameter.page) {
    const template = HtmlService.createTemplateFromFile('Latest Readings');
    return template.evaluate();
  }
  const template = HtmlService.createTemplateFromFile(e.parameter['page']);
  return template.evaluate();
}

function getAllSensorData() {
  const sheet = getSheet('Readings Database');
  const data = sheet.getDataRange().getValues();
  const result = {};

  data.forEach(row => {
    const sensorIndexRaw = row[0];
    if (sensorIndexRaw === '' || sensorIndexRaw === null || sensorIndexRaw === undefined) return;

    // Convert to 0-indexed key (assuming sheet sensors start at 1)
    const sensorIndex = Number(sensorIndexRaw) - 1;
    if (isNaN(sensorIndex) || sensorIndex < 0) return;

    if (!result[sensorIndex]) result[sensorIndex] = [];

    // Parse timestamp properly
    let timestamp = null;
    if (row[1] instanceof Date) {
      timestamp = row[1];
    } else if (typeof row[1] === 'string') {
      const parsed = Date.parse(row[1].trim());
      if (!isNaN(parsed)) timestamp = new Date(parsed);
    }
    if (!timestamp) timestamp = new Date(0);

    // Format values safely
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm");

    const tempRaw = safeNumber(row[2]);
    const humRaw = safeNumber(row[3]);
    const presRaw = (row[4] === "N/A" || row[4] === null || row[4] === undefined) ? null : safeNumber(row[4]);

    result[sensorIndex].push([
      formattedDate,
      timeStr,
      tempRaw === null ? "N/A" : `${tempRaw.toFixed(1)}°C`,
      humRaw === null ? "N/A" : `${(Number(humRaw.toFixed(1)) * 100).toFixed(1)}%`,
      presRaw === null ? "N/A" : `${presRaw.toFixed(1)}Pa`
    ]);
  });

  return result;
}
function safeNumber(value) {
  const num = Number(value);
  return isNaN(num) ? null : num;
}

function getLatestSensorData() {
  const sheet = getSheet('Readings Database');
  const data = sheet.getDataRange().getValues();
  const latest = {};

  data.forEach(row => {
    const sensorIndexRaw = row[0];
    if (sensorIndexRaw === '' || sensorIndexRaw === null || sensorIndexRaw === undefined) return;

    // Convert to 0-indexed key (if sensor numbering starts at 1)
    const sensorIndex = Number(sensorIndexRaw) - 1;
    if (isNaN(sensorIndex) || sensorIndex < 0) return;

    // Parse timestamp robustly
    let timestamp = null;
    if (row[1] instanceof Date) {
      timestamp = row[1];
    } else if (typeof row[1] === 'string') {
      const parsed = Date.parse(row[1].trim());
      if (!isNaN(parsed)) timestamp = new Date(parsed);
    }
    if (!timestamp) timestamp = new Date(0);

    // Only keep the most recent reading per sensor
    if (!latest[sensorIndex] || timestamp > latest[sensorIndex].timestamp) {
      const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd");
      const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm");

      const tempRaw = safeNumber(row[2]);
      const humRaw = safeNumber(row[3]);
      const presRaw = (row[4] === "N/A" || row[4] === null || row[4] === undefined) ? null : safeNumber(row[4]);

      latest[sensorIndex] = {
        timestamp,
        reading: [
          formattedDate,
          timeStr,
          tempRaw === null ? "N/A" : `${tempRaw.toFixed(1)}°C`,
          humRaw === null ? "N/A" : `${(Number(humRaw.toFixed(1)) * 100).toFixed(1)}%`,
          presRaw === null ? "N/A" : `${presRaw.toFixed(1)}Pa`
        ]
      };
    }
  });

  // Return a clean 0-indexed result map
  const result = {};
  Object.keys(latest).forEach(index => {
    result[index] = latest[index].reading;
  });

  return result;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}

function getAlertInfo() {
  const sheet = getSheet('Alert Info');
  const alertRange = sheet.getRange(2, 1, 11, 9);
  const alertInfo = alertRange.getValues();
  for (let i = 0; i < alertInfo.length; i++) {
    const emailArray = alertInfo[i][8].split(",");
    alertInfo[i][8] = emailArray;
  }
  return alertInfo;
}

function doPost(e) {

  try {

    // Parse JSON body from ESP32

    const data = JSON.parse(e.postData.contents);



    // (Optional) security key

    const SECRET_KEY = "abc123";

    if (data.key && data.key !== SECRET_KEY) {

      return ContentService.createTextOutput("Unauthorized");

    }



    // Open the active spreadsheet and sheet

    const sheet = getSheet("Readings Database");



    // Append new row with timestamp and data

    sheet.appendRow([

      data.unit || "",       // sensor number

      new Date(),              // Timestamp

      data.temperature || "",  // Temperature value

      data.humidity || "",    // Humidity value

      data.pressure || ""       // pressure value

    ]);



    // Send confirmation back to ESP32

    return ContentService

      .createTextOutput("Data added successfully")

      .setMimeType(ContentService.MimeType.TEXT);



  } catch (error) {

    // Handle JSON or other errors

    return ContentService

      .createTextOutput("Error: " + error)

      .setMimeType(ContentService.MimeType.TEXT);

  }

}

