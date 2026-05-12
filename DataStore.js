const REQUEST_CACHE = {};
const USER_LIST_CACHE_KEY = 'soar_users_table_v1';
const USER_LIST_CACHE_SECONDS = 300;
const ID_COUNTER_PROPERTY_PREFIX = 'soar_next_id:';
const DATA_VERSION_PROPERTY_KEY = 'soar_data_version:last_updated';
function clearRequestCache(sheetName) {
  if (sheetName) {
    delete REQUEST_CACHE[sheetName];
    delete REQUEST_CACHE[`${sheetName}:raw`];
    return;
  }
  Object.keys(REQUEST_CACHE).forEach((key) => delete REQUEST_CACHE[key]);
}
function bumpDataVersion() {
  const timestamp = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty(DATA_VERSION_PROPERTY_KEY, timestamp);
  return timestamp;
}
function getStoredDataVersion() {
  const properties = PropertiesService.getScriptProperties();
  const storedVersion = properties.getProperty(DATA_VERSION_PROPERTY_KEY);
  if (storedVersion) return storedVersion;
  return bumpDataVersion();
}
function getSpreadsheetLastUpdatedTimestamp() {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    return DriveApp.getFileById(spreadsheetId).getLastUpdated().toISOString();
  } catch (error) {
    Logger.log(`Falling back to script data version: ${error && error.message ? error.message : error}`);
    return getStoredDataVersion();
  }
}
function buildGlobalVersionHash() {
  const spreadsheetUpdatedAt = getSpreadsheetLastUpdatedTimestamp();
  const appUpdatedAt = getStoredDataVersion();
  return Utilities.base64EncodeWebSafe(`${spreadsheetUpdatedAt}|${appUpdatedAt}`);
}
function getGlobalVersionHash() {
  const spreadsheetUpdatedAt = getSpreadsheetLastUpdatedTimestamp();
  const appUpdatedAt = getStoredDataVersion();
  return JSON.stringify({
    success: true,
    versionHash: Utilities.base64EncodeWebSafe(`${spreadsheetUpdatedAt}|${appUpdatedAt}`),
    lastUpdated: appUpdatedAt,
    spreadsheetUpdatedAt: spreadsheetUpdatedAt,
    currentUserEmail: normalizeEmail(getCurrentUser())
  });
}
function invalidateTableCache(sheetName) {
  clearRequestCache(sheetName);
  if (sheetName === 'Users') {
    CacheService.getScriptCache().remove(USER_LIST_CACHE_KEY);
  }
  bumpDataVersion();
}
function rowsToObjects(headers, rows) {
  return rows.map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}
function normalizeValueForClient(value) {
  return value instanceof Date ? value.toISOString() : value;
}
function getCachedSheetValues(sheetName) {
  const cacheKey = `${sheetName}:raw`;
  if (REQUEST_CACHE[cacheKey]) return REQUEST_CACHE[cacheKey];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    REQUEST_CACHE[cacheKey] = { sheet: null, data: [] };
    return REQUEST_CACHE[cacheKey];
  }

  const data = sheet.getDataRange().getValues();
  REQUEST_CACHE[cacheKey] = { sheet, data };
  return REQUEST_CACHE[cacheKey];
}
function getTableData(sheetName) {
  if (REQUEST_CACHE[sheetName]) return REQUEST_CACHE[sheetName];

  if (sheetName === 'Users') {
    const cachedUsers = CacheService.getScriptCache().get(USER_LIST_CACHE_KEY);
    if (cachedUsers) {
      try {
        REQUEST_CACHE[sheetName] = JSON.parse(cachedUsers);
        return REQUEST_CACHE[sheetName];
      } catch (error) {
        CacheService.getScriptCache().remove(USER_LIST_CACHE_KEY);
      }
    }
  }

  const sheetValues = getCachedSheetValues(sheetName);
  const data = sheetValues.data;
  if (data.length <= 1) {
    REQUEST_CACHE[sheetName] = [];
    return REQUEST_CACHE[sheetName];
  }

  const tableData = rowsToObjects(data[0], data.slice(1));
  REQUEST_CACHE[sheetName] = tableData;

  if (sheetName === 'Users') {
    try {
      CacheService.getScriptCache().put(USER_LIST_CACHE_KEY, JSON.stringify(tableData), USER_LIST_CACHE_SECONDS);
    } catch (error) {
      Logger.log(`Skipping user cache write: ${error && error.message ? error.message : error}`);
    }
  }

  return tableData;
}
function getIdCounterPropertyKey(sheetName, prefix) {
  return `${ID_COUNTER_PROPERTY_PREFIX}${sheetName}:${prefix}`;
}
function seedIdCounterFromSheet(sheetName, prefix) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return 0;

  const firstColumnValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return firstColumnValues.reduce((max, row) => {
    const id = (row[0] || '').toString();
    if (!id.startsWith(`${prefix}-`)) return max;
    const numericPart = parseInt(id.split('-')[1], 10);
    return Number.isNaN(numericPart) ? max : Math.max(max, numericPart);
  }, 0);
}
function generateNextId(sheetName, prefix) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const properties = PropertiesService.getScriptProperties();
    const propertyKey = getIdCounterPropertyKey(sheetName, prefix);
    const currentValue = properties.getProperty(propertyKey);
    const currentNumber = currentValue ? parseInt(currentValue, 10) : seedIdCounterFromSheet(sheetName, prefix);
    const nextNumber = (Number.isNaN(currentNumber) ? 0 : currentNumber) + 1;
    properties.setProperty(propertyKey, nextNumber.toString());
    return `${prefix}-${nextNumber.toString().padStart(8, '0')}`;
  } finally {
    lock.releaseLock();
  }
}
function appendRows(sheet, rows) {
  if (!sheet || !rows || rows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
function deleteRowsBySheetIndexes(sheet, sheetRowIndexes) {
  if (!sheet || !sheetRowIndexes || sheetRowIndexes.length === 0) return;
  const sortedIndexes = [...new Set(sheetRowIndexes)].sort((a, b) => b - a);
  let groupStart = sortedIndexes[0];
  let groupCount = 1;

  for (let i = 1; i <= sortedIndexes.length; i++) {
    const current = sortedIndexes[i];
    if (current === groupStart - groupCount) {
      groupCount += 1;
      continue;
    }

    sheet.deleteRows(groupStart - groupCount + 1, groupCount);
    groupStart = current;
    groupCount = 1;
  }
}
function updateRowValues(sheet, sheetRowIndex, values) {
  if (!sheet || !sheetRowIndex || !values) return;
  sheet.getRange(sheetRowIndex, 1, 1, values.length).setValues([values]);
}
