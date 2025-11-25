const TEST_MODE = false; // set to true to enable test mode (no emails sent)

/** ---------- Utility Helpers ---------- **/

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function safeNumber(value) {
  const num = Number(value);
  return isNaN(num) ? null : num;
}

/** ---------- View Rendering ---------- **/

const PAGE_WHITELIST = ["Latest Readings", "Readings History"]; // add allowed files

function doGet(e) {
  const requested = (e && e.parameter && e.parameter.page) ? e.parameter.page : "Latest Readings";
  const page = PAGE_WHITELIST.includes(requested) ? requested : "Latest Readings";
  const template = HtmlService.createTemplateFromFile(page);
  template.sensor = e.parameter.sensor || null;
  console.log(">>> doGet SENSOR PARAM:", e.parameter.sensor);

  return template.evaluate().setTitle("Global Readings Project");
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/** ---------- Core Data Functions ---------- **/

/**
 * Loads and caches the entire readings database once per execution context.
 */
function getDatabaseCache_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("readings_db_cached");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      console.log("Cache parse failed, using backup cache");
    }
  }
  if (globalThis.__readingsCache) {
    return globalThis.__readingsCache;
  }

  const sheet = getSheet("Readings Database");
  const data = (sheet && sheet.getDataRange()) ? sheet.getDataRange().getValues() : [];
  try {
    cache.put("readings_db_cached", JSON.stringify(data), 600); // 600 seconds
  } catch (e) {
    // fallback to globalThis to avoid breaking behavior
    globalThis.__readingsCache = data;
  }
  return data;
}

function invalidateDatabaseCache() {
  const CACHE_KEY = "readings_db_cached";
  CacheService.getScriptCache().remove(CACHE_KEY);
  try { delete globalThis.__readingsCache; } catch (e) { /* ignore */ };
}

/**
 * Returns all sensor data, grouped by 0-indexed sensor ID.
 */
function getAllSensorData() {
  const data = getDatabaseCache_();
  console.log(data);
  const result = {};

  data.forEach(row => {
    const sensorRaw = row[0];
    if (!sensorRaw) return;

    const sensorIndex = Number(sensorRaw) - 1;
    if (isNaN(sensorIndex) || sensorIndex < 0) return;

    const timestamp = parseTimestamp(row[1]);
    if (!timestamp) {
      return;
    }
    const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss");

    const tempRaw = safeNumber(row[2]);
    const humRaw = safeNumber(row[3]);
    const presRaw = (row[4] === "N/A" || row[4] == null) ? null : safeNumber(row[4]);

    if (!result[sensorIndex]) result[sensorIndex] = [];
    result[sensorIndex].push([
      formattedDate,
      timeStr,
      tempRaw == null ? "N/A" : `${tempRaw.toFixed(1)}°C`,
      humRaw == null ? "N/A" : `${(humRaw * 100).toFixed(1)}%`,
      presRaw == null ? "N/A" : `${presRaw.toFixed(1)}Pa`
    ]);
  });
  console.log(result);
  return result;
}

/**
 * Returns only the latest reading per sensor (0-indexed).
 */
function getLatestSensorData() {
  const data = getDatabaseCache_();
  const latest = {};

  data.forEach(row => {
    const sensorRaw = row[0];
    if (!sensorRaw) return;

    const sensorIndex = Number(sensorRaw) - 1;
    if (isNaN(sensorIndex) || sensorIndex < 0) return;

    const timestamp = parseTimestamp(row[1]);
    if (!timestamp) {
      return;
    }
    if (!latest[sensorIndex] || timestamp > latest[sensorIndex].timestamp) {
      const formattedDate = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd");
      const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "HH:mm:ss");

      const tempRaw = safeNumber(row[2]);
      const humRaw = safeNumber(row[3]);
      const presRaw = (row[4] === "N/A" || row[4] == null) ? null : safeNumber(row[4]);

      latest[sensorIndex] = {
        timestamp,
        reading: [
          formattedDate,
          timeStr,
          tempRaw == null ? "N/A" : `${tempRaw.toFixed(1)}°C`,
          humRaw == null ? "N/A" : `${(humRaw * 100).toFixed(1)}%`,
          presRaw == null ? "N/A" : `${presRaw.toFixed(1)}Pa`
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

/**
 * Parses a timestamp value safely (supports Date or string).
 */
function parseTimestamp(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  if (typeof value === "string") {
    const parts = value.match(/(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2}):(\d{2})/);
    if (parts) {
      const [, dd, mm, yyyy, hh, min, ss] = parts.map(Number);
      return new Date(yyyy, mm - 1, dd, hh, min, ss);
    }
    // fallback for other formats
    const parsed = Date.parse(value.replace(/-/g, "/").trim());
    if (!isNaN(parsed)) return new Date(parsed);
  }

  if (typeof value === "number" && value > 0) {
    const ms = (value - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d.getTime())) return d;
  }

  return null;
}



/** ---------- Alerts Handling ---------- **/

let alertInfoCache = null;

function getAlertInfo() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("alert_info_cached");

  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) return parsed;
    } catch (e) { }
  }

  const sheet = getSheet("Alert Info");
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const numRows = lastRow - 1;
  const rows = sheet.getRange(2, 1, numRows, 9).getValues();

  // Clean email list (column 9)
  rows.forEach((row, i) => {
    if (typeof row[8] === "string") {
      rows[i][8] = row[8]
        .split(/[,;]/)
        .map(s => s.trim())
        .filter(Boolean);
    } else {
      rows[i][8] = [];
    }
  });

  cache.put("alert_info_cached", JSON.stringify(rows), 300);
  console.log(rows);
  return rows;
}


function invalidateAlertInfoCache() {
  alertInfoCache = null;
  try { CacheService.getScriptCache().remove("alert_info_cached"); } catch (e) { /* ignore */ }
}


/** ---------- Data Upload (ESP32) ---------- **/

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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Readings Database");

    let pressureVal = Number(data.pressure); // Pa (no conversion)
    if (isNaN(pressureVal)) pressureVal = "";

    // Append new row with timestamp and data
    sheet.appendRow([
      data.unit || "",         // sensor number
      new Date(),              // Timestamp
      data.temperature || "",  // Temperature value
      data.humidity || "",    // Humidity value
      pressureVal || ""      // pressure value
    ]);

    // --- Pass normalized data to alert system ---
    checkData({
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
      time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss"),
      temperature: isNaN(data.temperature) ? null : data.temperature,
      humidity: isNaN(data.humidity) ? null : data.humidity,
      pressure: isNaN(pressureVal) ? null : pressureVal,
      location: Number(data.unit) - 1
    });

    invalidateDatabaseCache();

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

/*function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const SECRET_KEY = "abc123";

    // require valid key explicitly
    if (!data.key || data.key !== SECRET_KEY) {
      return ContentService.createTextOutput("Unauthorized");
    }


    const sheet = getSheet("Readings Database");
    // --- Normalize values before saving ---
    const tempVal = Number(data.temperature);

    let humidityVal = Number(data.humidity);
    if (!isNaN(humidityVal) && humidityVal > 1) {
      // Convert percent to fraction (e.g. 85 -> 0.85)
      humidityVal = humidityVal / 100;
    }

    let pressureVal = Number(data.pressure); // Pa (no conversion)
    if (isNaN(pressureVal)) pressureVal = "";

    // --- Write to sheet ---
    sheet.appendRow([
      data.unit || "",
      new Date(),
      isNaN(tempVal) ? "" : tempVal,
      isNaN(humidityVal) ? "" : humidityVal,
      pressureVal
    ]);

    // --- Pass normalized data to alert system ---
    checkData({
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
      time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss"),
      temperature: isNaN(tempVal) ? null : tempVal,
      humidity: isNaN(humidityVal) ? null : humidityVal,
      pressure: isNaN(pressureVal) ? null : pressureVal,
      location: Number(data.unit) - 1
    });

    invalidateDatabaseCache();


    return ContentService
      .createTextOutput("Data added successfully")
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService
      .createTextOutput("Error: " + error)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}*/

function getLocations() {
  try {
    const alertInfo = getAlertInfo();

    console.log("RAW alertInfo rows:", JSON.stringify(alertInfo));

    // Extract column A values
    const locations = alertInfo
      .map(row => row[0])
      .filter(v => typeof v === "string" && v.trim() !== "");

    console.log("RETURNING LOCATIONS:", locations);
    return locations;

  } catch (err) {
    console.log(`getLocations error: ${err.message || err}`);
    return [];
  }
}

// Remove previous module-level alert arrays

function getActiveAlerts() {
  const props = PropertiesService.getScriptProperties();
  const json = props.getProperty("activeAlerts");
  if (json) return JSON.parse(json);
  return { temperature: [], humidity: [], pressure: [] };
}

function setActiveAlerts(active) {
  PropertiesService.getScriptProperties().setProperty("activeAlerts", JSON.stringify(active));
}

function checkData(newReading) {
  if (!newReading || newReading.location == null || isNaN(newReading.location)) {
    return;
  }

  const alertInfo = getAlertInfo();
  const active = getActiveAlerts();

  for (let i = 0; i < alertInfo.length; i++) {
    if (newReading.location !== i) continue;

    while (active.temperature.length <= i) active.temperature.push(false);
    while (active.humidity.length <= i) active.humidity.push(false);
    while (active.pressure.length <= i) active.pressure.push(false);

    const row = alertInfo[i];
    handleAlertsForRow(i, newReading, row, active);
  }

  setActiveAlerts(active);
}


function handleAlertsForRow(i, newReading, row, active) {
  const [
    locationNameRaw,
    tempHighRaw, tempLowRaw,
    humHighRaw, humLowRaw,
    presHighRaw, presLowRaw,
    , recipientsRaw
  ] = row;

  const locationName = locationNameRaw;
  const tempHigh = Number(tempHighRaw);
  const tempLow = Number(tempLowRaw);
  const humHigh = Number(humHighRaw);
  const humLow = Number(humLowRaw);
  const presHigh = Number(presHighRaw);
  const presLow = Number(presLowRaw);

  const recipients = Array.isArray(recipientsRaw)
    ? recipientsRaw
    : (typeof recipientsRaw === "string"
      ? recipientsRaw.split(/[;,]/).map(s => s.trim()).filter(Boolean)
      : []);

  const { temperature: temp, humidity: hum, pressure: pres } = newReading;

  const checks = [
    {
      key: "temperature",
      value: temp,
      high: tempHigh,
      low: tempLow,
      activeArray: active.temperature,
      alertFn: temperatureAlert,
      clearFn: clearTemperatureAlert,
    },
    {
      key: "humidity",
      value: hum,
      high: humHigh,
      low: humLow,
      activeArray: active.humidity,
      alertFn: humidityAlert,
      clearFn: clearHumidityAlert,
    },
    {
      key: "pressure",
      value: pres,
      high: presHigh,
      low: presLow,
      activeArray: active.pressure,
      alertFn: pressureAlert,
      clearFn: clearPressureAlert,
    },
  ];

  for (const { key, value, high, low, activeArray, alertFn, clearFn } of checks) {
    if (value == null || value === "" || isNaN(value)) continue;

    if (!isNaN(value) && !isNaN(high) && !isNaN(low) && (value >= high || value <= low)) {
      if (!activeArray[i]) {
        activeArray[i] = true;
        alertFn(locationName, recipients, newReading);
      }
    } else if (activeArray[i]) {
      activeArray[i] = false;
      clearFn(locationName, recipients, newReading);
    } else {
    }
  }
}



/** ---------- Safe Email Helper ---------- **/
function safeSendEmail(recipients, subject, body) {
  if (TEST_MODE) {
    console.log(`[TEST_MODE] Email suppressed → To: ${recipients.join(", ")} | Subject: ${subject}`);
    return;
  }
  try {
    MailApp.sendEmail({
      to: recipients.join(","),
      subject,
      body,
    });
  } catch (err) {
    console.log(`Email send failed for ${recipients}: ${err.message}`);
  }
}


/** ---------- Individual Alert Handlers ---------- **/

//Temperature alert triggered
function temperatureAlert(locationName, recipients, newReading) {
  const subject = `Abnormal Temperature Reading Detected`;
  const body = `Date: ${newReading.date}
Time: ${newReading.time}
Location: ${locationName}
Reading: ${newReading.temperature}°C`;

  safeSendEmail(recipients, subject, body);
}

//Temperature alert cleared
function clearTemperatureAlert(locationName, recipients, newReading) {
  const subject = `Abnormal Temperature Reading No Longer Detected`;
  const body = `Date: ${newReading.date}
Time: ${newReading.time}
Location: ${locationName}
Reading: ${newReading.temperature}°C`;

  safeSendEmail(recipients, subject, body);
}

//Humidity alert triggered
function humidityAlert(locationName, recipients, newReading) {
  const subject = `Abnormal Humidity Reading Detected`;
  const body = `Date: ${newReading.date}
Time: ${newReading.time}
Location: ${locationName}
Reading: ${(newReading.humidity * 100).toFixed(1)}%`;

  safeSendEmail(recipients, subject, body);
}

//Humidity alert cleared
function clearHumidityAlert(locationName, recipients, newReading) {
  const subject = `Abnormal Humidity Reading No Longer Detected`;
  const body = `Date: ${newReading.date}
Time: ${newReading.time}
Location: ${locationName}
Reading: ${(newReading.humidity * 100).toFixed(1)}%`;

  safeSendEmail(recipients, subject, body);
}

//Pressure alert triggered
function pressureAlert(locationName, recipients, newReading) {
  const subject = `Abnormal Pressure Reading Detected`;
  const body = `Date: ${newReading.date}
Time: ${newReading.time}
Location: ${locationName}
Reading: ${newReading.pressure} Pa`;

  safeSendEmail(recipients, subject, body);
}

//Pressure alert cleared
function clearPressureAlert(locationName, recipients, newReading) {
  const subject = `Abnormal Pressure Reading No Longer Detected`;
  const body = `Date: ${newReading.date}
Time: ${newReading.time}
Location: ${locationName}
Reading: ${newReading.pressure} Pa`;

  safeSendEmail(recipients, subject, body);
}

function testAlertSystem() {
  console.log("=== Running testAlertSystem ===");

  // Local test mode override (no real emails sent)
  const localTestMode = true;

  // Wrap safeSendEmail temporarily
  const originalSafeSendEmail = safeSendEmail;
  safeSendEmail = function (recipients, subject, body) {
    if (localTestMode) {
      console.log(`[TEST_MODE] Email suppressed → To: ${recipients.join(", ")} | Subject: ${subject}`);
      console.log("Email body preview:\n" + body);
      return;
    }
    originalSafeSendEmail(recipients, subject, body);
  };

  invalidateAlertInfoCache();

  const fakeTriggerReading = {
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
    time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss"),
    temperature: 18,
    humidity: 0.50,
    pressure: 110,
    location: 4,
  };

  console.log("→ Simulating abnormal reading to TRIGGER alert");
  checkData(fakeTriggerReading);

  Utilities.sleep(1000);

  const fakeClearReading = {
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd"),
    time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss"),
    temperature: 30,
    humidity: 0.81,
    pressure: 111,
    location: 4,
  };

  console.log("→ Simulating normal reading to CLEAR alert");
  checkData(fakeClearReading);

  console.log("=== Test complete ===");

  // Restore original email sender
  safeSendEmail = originalSafeSendEmail;
}