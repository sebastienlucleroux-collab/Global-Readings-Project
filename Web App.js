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
      Logger.log("Cache parse failed, using backup cache");
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
  const result = {};

  data.forEach(row => {
    const sensorRaw = row[0];
    if (!sensorRaw) return;

    const sensorIndex = Number(sensorRaw) - 1;
    if (isNaN(sensorIndex) || sensorIndex < 0) return;

    const timestamp = parseTimestamp_(row[1]);
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

    const timestamp = parseTimestamp_(row[1]);
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
function parseTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "string") {
    const parsed = Date.parse(value.replace(/-/g, "/").trim());
    if (!isNaN(parsed)) return new Date(parsed);
  }
  return null; // explicit invalid timestamp
}


/** ---------- Alerts Handling ---------- **/

function getAlertInfo() {
  const sheet = getSheet("Alert Info");
  const alertRange = sheet.getRange(2, 1, 11, 9).getValues();

  alertRange.forEach((row, i) => {
    if (typeof row[8] === "string") {
      alertRange[i][8] = row[8].split(",").map(email => email.trim());
    }
  });

  return alertRange;
}

/** ---------- Data Upload (ESP32) ---------- **/

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const SECRET_KEY = "abc123";

    if (data.key && data.key !== SECRET_KEY) {
      return ContentService.createTextOutput("Unauthorized");
    }

    const sheet = getSheet("Readings Database");
    sheet.appendRow([
      data.unit || "",
      new Date(),
      data.temperature || "",
      data.humidity || "",
      data.pressure || ""
    ]);

    // Invalidate cache so next load is fresh
    invalidateDatabaseCache();

    return ContentService
      .createTextOutput("Data added successfully")
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService
      .createTextOutput("Error: " + error)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function getLocations() {
  try {
    const alertInfo = getAlertInfo();
    return alertInfo.map(row => row[0]).filter(Boolean);
  } catch (err) {
    Logger.log(`getLocations error: ${err.message || err}`);
    return [];
  }
}