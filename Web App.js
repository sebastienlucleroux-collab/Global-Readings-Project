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
    const sensorIndex = row[0];
    if (!result[sensorIndex]) result[sensorIndex] = [];

    row.shift();
    const [dateStr, timeStr] = row[0].split(' ');
    const [month, day, year] = dateStr.split('/');
    const formattedDate = `${year}/${month.padStart(2, '0')}/${day.padStart(2, '0')}`;

    result[sensorIndex].push([
      formattedDate,
      timeStr,
      `${row[1].toFixed(1)}°C`,
      `${(Number(row[2].toFixed(1)) * 100).toFixed(1)}%`,
      row[3] === "N/A" ? row[3] : `${row[3].toFixed(1)}Pa`
    ]);
  });

  return result;
}

function getLatestSensorData() {
  const sheet = getSheet('Readings Database');
  const data = sheet.getDataRange().getValues();
  const latest = {};

  data.forEach(row => {
    const sensorIndex = row[0];
    const timestamp = new Date(row[1]);

    if (!latest[sensorIndex] || timestamp > latest[sensorIndex].timestamp) {
      const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd");
      const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm");

      latest[sensorIndex] = {
        timestamp,
        reading: [
          formattedDate,
          timeStr,
          `${row[2].toFixed(1)}°C`,
          `${(Number(row[3].toFixed(1)) * 100).toFixed(1)}%`,
          row[4] === "N/A" ? row[4] : `${row[4].toFixed(1)}Pa`
        ]
      };
    }
  });

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

