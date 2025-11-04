//NB: All iterations of 'testDataArray' must be replaced

//List of all locations
const locationArray = ["Pressure1", "Pressure2", "Pressure3", "Pressure4", "Pressure5", "NoPressure1", "NoPressure2", "NoPressure3", "NoPressure4", "NoPressure5", "NoPressure6"];

//number of 'false' = Total number of locations
var temperatureAlertActive = [false, false, false, false, false, false, false, false, false, false, false];
var humidityAlertActive = [false, false, false, false, false, false, false, false, false, false, false];
var pressureAlertActive = [false, false, false, false, false, false, false, false, false, false, false];

function fetchData(){
  // testDataArray = logged readings
  const testDataArray = ["2022/07/26", "10:00:00", "19", "0.37", "48"];
  return testDataArray;
}

function writeData() {
  const testDataArray = fetchData()
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2];
  const lastRow = sheet.getLastRow();
  newRowNumber = lastRow+1;
  for(let i=0; i<5; i++) {
    const cell = sheet.getRange(newRowNumber, (i+1));
    cell.setValue(testDataArray[i]);
  }
  checkData();
}

function checkData() {

var temperatureAlertActive = [false, false, false, false, false, false, false, false, false, false, false];
var humidityAlertActive = [false, false, false, false, false, false, false, false, false, false, false];
var pressureAlertActive = [false, false, false, false, false, false, false, false, false, false, false];

  let alertInfo = getAlertInfo()
  //i < Total number of locations
  for (let i = 0; i < 11; i++) {
    const testDataArray = fetchData();
    if (JSON.parse(testDataArray[2]) >= alertInfo[i][1] || JSON.parse(testDataArray[2]) <= alertInfo[i][2]) {
      if(!temperatureAlertActive[i]) {
        temperatureAlertActive[i] = true;
        temperatureAlert(i);
      }
    }else{
      if(temperatureAlertActive[i]) {
        temperatureAlertActive[i] = false;        
        clearTemperatureAlert(i);
      }
    }

    if (JSON.parse(testDataArray[3]) >= alertInfo[i][3] || JSON.parse(testDataArray[3]) <= alertInfo[i][4]) {
      if(!humidityAlertActive[i]) {
        humidityAlertActive[i] = true;
        humidityAlert(i);
      }
    }else{
      if(humidityAlertActive[i]) {
        humidityAlertActive[i] = false;
        clearHumidityAlert(i);
      }
    }

    if(testDataArray[4]) {
      if (JSON.parse(testDataArray[4]) >= alertInfo[i][5] || JSON.parse(testDataArray[4]) <= alertInfo[i][6]) {
        if(!pressureAlertActive[i]) {
          pressureAlertActive[i] = true;
          pressureAlert(i);
        }
      }else{
        if(pressureAlertActive[i]) {
          pressureAlertActive[i] = false;
          clearPressureAlert(i);
        }
      }
    }
  }
}

function getUnformattedAlertInfo(){

  
  //13 = index of alert info sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[13];
  //11 = Total number of locations
  const alertRange = sheet.getRange(2, 1, 11, 9);
  const alertInfo = alertRange.getValues();
  //i < Total number of locations

  return alertInfo;
}

function getAlertInfo(){
  const alertInfo = getUnformattedAlertInfo();
  for (let i = 0; i < 11; i++){
    const emailArray = alertInfo[i][8].split(";");
    alertInfo[i][8] = emailArray;
  }
  return alertInfo;
}

function temperatureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  for (let j = 0; j < alertInfo[i][8].length; j++) {
    var recipient = alertInfo[i][8][j];
    var subject = `Abnormal Temperature Reading Detected`;
    var body = `Date: ${testDataArray[0]}
Time: ${testDataArray[1]}
Location: ${alertInfo[i][0]}
Reading: ${testDataArray[2]}°C`;
    MailApp.sendEmail(recipient, subject, body);
  }
}

function clearTemperatureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  for (let j = 0; j < alertInfo[i][8].length; j++) {
    var recipient = alertInfo[i][8][j];
    var subject = `Abnormal Temperature Reading No Longer Detected`;
    var body = `Date: ${testDataArray[0]}
Time: ${testDataArray[1]}
Location: ${alertInfo[i][0]}
Reading: ${testDataArray[2]}°C`;
    MailApp.sendEmail(recipient, subject, body);
  }
}


function humidityAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  for (let j = 0; j < alertInfo[i][8].length; j++) {
    var recipient = alertInfo[i][8][j];
    var subject = `Abnormal Humidity Reading Detected`;
    var body = `Date: ${testDataArray[0]}
Time: ${testDataArray[1]}
Location: ${alertInfo[i][0]}
Reading: ${(testDataArray[3])*100}%`;
    MailApp.sendEmail(recipient, subject, body);
  }
}

function clearHumidityAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  for (let j = 0; j < alertInfo[i][8].length; j++) {
    var recipient = alertInfo[i][8][j];
    var subject = `Abnormal Humidity Reading No Longer Detected`;
    var body =`Date: ${testDataArray[0]}
Time: ${testDataArray[1]}
Location: ${alertInfo[i][0]}
Reading: ${(testDataArray[3])*100}%`;
    MailApp.sendEmail(recipient, subject, body);
  }
}

function pressureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  for (let j = 0; j < alertInfo[i][8].length; j++) {
    var recipient = alertInfo[i][8][j];
    var subject = `Abnormal Pressure Reading Detected`;
    var body = `Date: ${testDataArray[0]}
Time: ${testDataArray[1]}
Location: ${alertInfo[i][0]}
Reading: ${testDataArray[4]}kPa`;
    MailApp.sendEmail(recipient, subject, body);
  }
}

function clearPressureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  for (let j = 0; j < alertInfo[i][8].length; j++) {
    var recipient = alertInfo[i][8][j];
    var subject = `Abnormal Pressure Reading No Longer Detected`;
    var body = `Date: ${testDataArray[0]}
Time: ${testDataArray[1]}
Location: ${alertInfo[i][0]}
Reading: ${testDataArray[4]}kPa`;
    MailApp.sendEmail(recipient, subject, body);
  }
}

/*
function temperatureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  var recipient = alertInfo[i][8];
  var subject = `Abnormal Temperature Reading Detected`;
  var body = `Date: ${testDataArray[0]}
  Time: ${testDataArray[1]}
  Location: ${alertInfo[i][0]}
  Reading: ${testDataArray[2]}°C`;
  MailApp.sendEmail(recipient, subject, body);
}

function clearTemperatureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  var recipient = alertInfo[i][8];
  var subject = `Abnormal Temperature Reading No Longer Detected`;
  var body = `Date: ${testDataArray[0]}
  Time: ${testDataArray[1]}
  Location: ${alertInfo[i][0]}
  Reading: ${testDataArray[2]}°C`;
  MailApp.sendEmail(recipient, subject, body);
}

function humidityAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  var recipient = alertInfo[i][8];
  var subject = `Abnormal Humidity Reading Detected`;
  var body = `Date: ${testDataArray[0]}
  Time: ${testDataArray[1]}
  Location: ${alertInfo[i][0]}
  Reading: ${(testDataArray[3])*100}%`;
  MailApp.sendEmail(recipient, subject, body);
}

function clearHumidityAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  var recipient = alertInfo[i][8];
  var subject = `Abnormal Humidity Reading No Longer Detected`;
  var body = `Date: ${testDataArray[0]}
  Time: ${testDataArray[1]}
  Location: ${alertInfo[i][0]}
  Reading: ${(testDataArray[3])*100}%`;
  MailApp.sendEmail(recipient, subject, body);
}

function pressureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  var recipient = alertInfo[i][8];
  var subject = `Abnormal Pressure Reading Detected`;
  var body = `Date: ${testDataArray[0]}
  Time: ${testDataArray[1]}
  Location: ${alertInfo[i][0]}
  Reading: ${testDataArray[4]}kPa`;
  MailApp.sendEmail(recipient, subject, body);
}

function clearPressureAlert(i) {
  const alertInfo = getAlertInfo();
  const testDataArray = fetchData();
  var recipient = alertInfo[i][8];
  var subject = `Abnormal Pressure Reading No Longer Detected`;
  var body = `Date: ${testDataArray[0]}
  Time: ${testDataArray[1]}
  Location: ${alertInfo[i][0]}
  Reading: ${testDataArray[4]}kPa`;
  MailApp.sendEmail(recipient, subject, body);
}
*/