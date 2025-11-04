function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function formatReadingRow(row) {
  row.shift();
  const [dateStr, timeStr] = row[0].split(' ');
  const [month, day, year] = dateStr.split('/');
  return {
    date: `${year}/${month.padStart(2, '0')}/${day.padStart(2, '0')}`,
    time: timeStr,
    temperature: `${row[1].toFixed(1)}°C`,
    humidity: `${(Number(row[2].toFixed(1)) * 100).toFixed(1)}%`,
    pressure: row[3] === "N/A" ? row[3] : `${row[3].toFixed(1)}Pa`
  };
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

function getSensorData(sensorIndex) {
  const sheet = getSheet('Readings Database');
  const data = sheet.getDataRange().getValues();
  const numericIndex = Number(sensorIndex);

  const sensorData = data
    .filter(row => row[0] === numericIndex)
    .map(row => {
      row.shift();

      const [dateStr, timeStr] = row[0].split(' ');
      const [month, day, year] = dateStr.split('/');
      const formattedDate = `${year}/${month.padStart(2, '0')}/${day.padStart(2, '0')}`;

      const temperature = `${row[1].toFixed(1)}°C`;
      const humidity = `${(Number(row[2].toFixed(1)) * 100).toFixed(1)}%`;
      const pressure = (row[3] === "N/A") ? row[3] : `${row[3].toFixed(1)}Pa`;

      return [formattedDate, timeStr, temperature, humidity, pressure];
    });

  return sensorData;
}

/*
function getSensorData(sensorIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Readings Database');
  const data = sheet.getDataRange().getValues();
  const sensorData = [];

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == Number(sensorIndex)) {
      data[i].shift()

      const dateTimeParts = data[i][0].split(' ');
      const dateParts = dateTimeParts[0].split('/');
      data[i][0] = [dateParts[2], dateParts[0].padStart(2,'0'), dateParts[1].padStart(2,'0')].join('/');
      data[i].splice(1, 0, dateTimeParts[1]);

      data[i][2] = `${(data[i][2]).toFixed(1)}°C`
      data[i][3] = `${(Number(data[i][3].toFixed(1))*100).toFixed(1)}%`
      data[i][4] = (data[i][4] === "N/A") ? data[i][4] : `${(data[i][4]).toFixed(1)}Pa`;

      sensorData.push(data[i]);
    }
  }

  return sensorData;
}
*/

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

