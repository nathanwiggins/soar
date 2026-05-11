/**
 * Main entry point for the web app.
 * Serves the Index.html file.
 *
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Soar - Project Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Includes HTML files into the main Index file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets the current user's email.
 */
function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}


function getCurrentUserProfilePhotoUrl() {
  const activeEmail = normalizeEmail(getCurrentUser());
  if (!activeEmail) return '';

  try {
    const response = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return '';
    }

    const data = JSON.parse(response.getContentText() || '{}');
    const profileEmail = normalizeEmail(data.email);
    if (profileEmail && profileEmail !== activeEmail) {
      return '';
    }

    return data.picture ? data.picture.toString().trim() : '';
  } catch (error) {
    return '';
  }
}

function syncCurrentUserProfilePhoto(sheet, data, headerIndex, currentUserEmail, profilePicUrl) {
  if (!sheet || !data || !headerIndex) return false;
  if (!currentUserEmail || !profilePicUrl) return false;

  const emailColumnIndex = headerIndex.Email;
  const profilePicColumnIndex = headerIndex.Profile_Pic_Url;

  if (emailColumnIndex === undefined || profilePicColumnIndex === undefined) {
    return false;
  }

  const matchingRowIndex = data.findIndex(
    (row, index) => index > 0 && normalizeEmail(row[emailColumnIndex]) === currentUserEmail
  );

  if (matchingRowIndex < 0) {
    return false;
  }

  const existingProfilePicUrl = data[matchingRowIndex][profilePicColumnIndex]
    ? data[matchingRowIndex][profilePicColumnIndex].toString().trim()
    : '';

  if (existingProfilePicUrl === profilePicUrl) {
    return false;
  }

  try {
    sheet.getRange(matchingRowIndex + 1, profilePicColumnIndex + 1).setValue(profilePicUrl);
    return true;
  } catch (error) {
    // Do not block app load for users who cannot edit the spreadsheet.
    return false;
  }
}

function normalizeEmail(email) {
  return email ? email.toString().trim().toLowerCase() : '';
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getHeaderIndex(headers) {
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}

function normalizeHeaderName(header) {
  return header ? header.toString().trim().toLowerCase() : '';
}

function getNormalizedHeaderIndex(headers) {
  return headers.reduce((acc, header, index) => {
    acc[normalizeHeaderName(header)] = index;
    return acc;
  }, {});
}

function getUserEmailFromRow(row, headerIndex) {
  if (!row || !headerIndex) return '';
  const emailColumnIndex = headerIndex.email;
  if (emailColumnIndex === undefined) return '';
  return normalizeEmail(row[emailColumnIndex]);
}

function safeSendEmail(recipient, subject, body) {
  const normalizedRecipient = normalizeEmail(recipient);
  if (!normalizedRecipient) return false;

  try {
    MailApp.sendEmail(normalizedRecipient, subject || '', body || '');
    return true;
  } catch (error) {
    // Notifications should never block core CRUD actions.
    Logger.log(`Failed to send notification email to ${normalizedRecipient}: ${error && error.message ? error.message : error}`);
    return false;
  }
}


const USER_SETTINGS_PROPERTY_PREFIX = 'soar_user_settings:';
const DEFAULT_NOTIFICATION_SETTINGS = {
  taskAssignments: true,
  commentsAndMentions: true,
  dueDateReminders: true,
  weeklyDigest: false
};

function getDefaultNotificationSettings() {
  return Object.assign({}, DEFAULT_NOTIFICATION_SETTINGS);
}

function normalizeNotificationSettings(settingsInput) {
  const normalized = getDefaultNotificationSettings();
  const source = settingsInput && typeof settingsInput === 'object' ? settingsInput : {};
  Object.keys(normalized).forEach((key) => {
    if (typeof source[key] === 'boolean') {
      normalized[key] = source[key];
    }
  });
  return normalized;
}

function getUserSettingsPropertyKey(email) {
  const normalizedEmail = normalizeEmail(email);
  return normalizedEmail ? `${USER_SETTINGS_PROPERTY_PREFIX}${normalizedEmail}` : '';
}

function getUserSettingsByEmail(email) {
  const propertyKey = getUserSettingsPropertyKey(email);
  if (!propertyKey) {
    return { notifications: getDefaultNotificationSettings() };
  }

  const rawSettings = PropertiesService.getScriptProperties().getProperty(propertyKey);
  if (!rawSettings) {
    return { notifications: getDefaultNotificationSettings() };
  }

  try {
    const parsedSettings = JSON.parse(rawSettings);
    return {
      notifications: normalizeNotificationSettings(parsedSettings && parsedSettings.notifications)
    };
  } catch (error) {
    Logger.log(`Failed to parse notification settings for ${normalizeEmail(email)}: ${error && error.message ? error.message : error}`);
    return { notifications: getDefaultNotificationSettings() };
  }
}

function isNotificationEnabledForEmail(email, notificationKey) {
  const defaults = getDefaultNotificationSettings();
  if (!Object.prototype.hasOwnProperty.call(defaults, notificationKey)) return true;

  const settings = getUserSettingsByEmail(email);
  return settings.notifications[notificationKey] !== false;
}

function isNotificationEnabledForUser(user, notificationKey) {
  return isNotificationEnabledForEmail(user && user.Email, notificationKey);
}

function persistCurrentUserSettings(settingsInput) {
  const currentUserEmail = normalizeEmail(getCurrentUser());
  if (!currentUserEmail) {
    return JSON.stringify({ success: false, error: 'Could not determine current user email.' });
  }

  const existingSettings = getUserSettingsByEmail(currentUserEmail);
  const requestedSettings = settingsInput && typeof settingsInput === 'object' ? settingsInput : {};
  const settingsToSave = {
    notifications: normalizeNotificationSettings(
      Object.prototype.hasOwnProperty.call(requestedSettings, 'notifications')
        ? requestedSettings.notifications
        : existingSettings.notifications
    )
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(getUserSettingsPropertyKey(currentUserEmail), JSON.stringify(settingsToSave));

  return JSON.stringify({ success: true, settings: settingsToSave });
}

function parseDateInput(dateInput) {
  if (!dateInput) return '';

  if (Object.prototype.toString.call(dateInput) === '[object Date]') {
    if (Number.isNaN(dateInput.getTime())) return '';
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }

  const trimmedValue = dateInput.toString().trim();
  if (!trimmedValue) return '';

  const dateOnlyMatch = trimmedValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateOnlyMatch) {
    const year = Number(dateOnlyMatch[1]);
    const month = Number(dateOnlyMatch[2]) - 1;
    const day = Number(dateOnlyMatch[3]);
    return new Date(year, month, day);
  }

  const parsedDate = new Date(trimmedValue);
  if (parsedDate.toString() === 'Invalid Date') return '';
  return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
}

function serializeDateOnlyForClient(dateValue) {
  if (!(Object.prototype.toString.call(dateValue) === '[object Date]') || Number.isNaN(dateValue.getTime())) {
    return dateValue;
  }
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

const TASK_COMPLETION_METADATA_SHEET = 'TaskCompletionMetadata';

function getTaskCompletionSheet(createIfMissing) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(TASK_COMPLETION_METADATA_SHEET);
  if (!sheet && createIfMissing) {
    sheet = spreadsheet.insertSheet(TASK_COMPLETION_METADATA_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['Task_ID', 'Completed_By', 'Completed_At']]);
  }
  return sheet;
}

function getTaskCompletionMetadataMap() {
  const sheet = getTaskCompletionSheet(false);
  const byTaskId = new Map();
  if (!sheet) return byTaskId;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return byTaskId;
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);
  const taskIdIndex = headerIndex.Task_ID;
  const completedByIndex = headerIndex.Completed_By;
  const completedAtIndex = headerIndex.Completed_At;
  if (taskIdIndex === undefined) return byTaskId;

  data.slice(1).forEach((row) => {
    const taskId = (row[taskIdIndex] || '').toString().trim();
    if (!taskId) return;
    const completedAtValue = completedAtIndex === undefined ? '' : row[completedAtIndex];
    byTaskId.set(taskId, {
      Completed_By: completedByIndex === undefined ? '' : (row[completedByIndex] || '').toString().trim(),
      Completed_At: completedAtValue instanceof Date ? completedAtValue.toISOString() : completedAtValue
    });
  });

  return byTaskId;
}

function upsertTaskCompletionMetadata(taskId, completedBy, completedAt) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) return;
  const sheet = getTaskCompletionSheet(true);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const headerIndex = getHeaderIndex(headers);
  const taskIdIndex = headerIndex.Task_ID;
  const completedByIndex = headerIndex.Completed_By;
  const completedAtIndex = headerIndex.Completed_At;
  if (taskIdIndex === undefined || completedByIndex === undefined || completedAtIndex === undefined) return;

  const rowIndex = data.findIndex((row, index) => index > 0 && (row[taskIdIndex] || '').toString().trim() === normalizedTaskId);
  const normalizedCompletedBy = completedBy ? completedBy.toString().trim() : '';
  const completedAtValue = completedAt || new Date();
  if (rowIndex > -1) {
    const updatedRow = data[rowIndex].slice();
    updatedRow[completedByIndex] = normalizedCompletedBy;
    updatedRow[completedAtIndex] = completedAtValue;
    updateRowValues(sheet, rowIndex + 1, updatedRow);
    return;
  }

  const newRow = new Array(headers.length).fill('');
  newRow[taskIdIndex] = normalizedTaskId;
  newRow[completedByIndex] = normalizedCompletedBy;
  newRow[completedAtIndex] = completedAtValue;
  appendRows(sheet, [newRow]);
}

function removeTaskCompletionMetadata(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) return;
  const sheet = getTaskCompletionSheet(false);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const headerIndex = getHeaderIndex(data[0]);
  const taskIdIndex = headerIndex.Task_ID;
  if (taskIdIndex === undefined) return;

  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][taskIdIndex] || '').toString().trim() === normalizedTaskId) {
      rowsToDelete.push(i + 1);
    }
  }
  deleteRowsBySheetIndexes(sheet, rowsToDelete);
}

function appendTaskCompletionMetadataToTask(task, metadataByTaskId) {
  if (!task) return task;
  const taskId = (task.Task_ID || '').toString().trim();
  if (!taskId) return task;
  const metadataMap = metadataByTaskId || getTaskCompletionMetadataMap();
  const metadata = metadataMap.get(taskId);

  const completedBy = (task.Completed_By || '').toString().trim();
  const completedAt = task.Completed_At || '';
  if (completedBy || completedAt) return task;

  if (metadata) {
    task.Completed_By = metadata.Completed_By || '';
    task.Completed_At = metadata.Completed_At || '';
  }
  return task;
}

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

/**
 * Core DB Function: Reads a sheet and returns an array of JSON objects.
 * Uses a per-execution request cache to avoid duplicate Spreadsheet reads.
 */
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

/**
 * Generates the zero-padded ID defined in the README using PropertiesService counters.
 */
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

function addUser(userInput) {
  const normalizedEmail = normalizeEmail(userInput && userInput.email);
  const name = userInput && userInput.name ? userInput.name.toString().trim() : '';
  const managerId = userInput && userInput.managerId ? userInput.managerId.toString().trim() : '';
  const currentUserEmail = normalizeEmail(getCurrentUser());
  const profilePicUrl = normalizedEmail === currentUserEmail ? getCurrentUserProfilePhotoUrl() : '';

  if (!normalizedEmail) {
    return JSON.stringify({ success: false, error: 'Email is required.' });
  }

  if (!isValidEmail(normalizedEmail)) {
    return JSON.stringify({ success: false, error: 'Email format is invalid.' });
  }

  if (!name) {
    return JSON.stringify({ success: false, error: 'Name is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Users sheet was not found.' });
  }

  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const headerIndex = getHeaderIndex(headers);
    const emailColumnIndex = headerIndex.Email;
    const userIdColumnIndex = headerIndex.User_ID;

    if (emailColumnIndex === undefined || userIdColumnIndex === undefined) {
      throw new Error('Users sheet must contain User_ID and Email columns.');
    }

    const existingUser = data
      .slice(1)
      .find((row) => normalizeEmail(row[emailColumnIndex]) === normalizedEmail);

    if (existingUser) {
      const user = {};
      headers.forEach((header, index) => {
        user[header] = existingUser[index];
      });
      return JSON.stringify({ success: true, user: user, created: false });
    }

    if (managerId && headerIndex.Manager_ID !== undefined) {
      const users = getTableData('Users');
      const validManagerIds = new Set(users.map((user) => user.User_ID));
      if (!validManagerIds.has(managerId)) {
        throw new Error(`Manager_ID ${managerId} does not exist.`);
      }
    }

    const newRow = new Array(headers.length).fill('');
    if (headerIndex.User_ID !== undefined) newRow[headerIndex.User_ID] = generateNextId('Users', 'U');
    if (headerIndex.Email !== undefined) newRow[headerIndex.Email] = normalizedEmail;
    if (headerIndex.Name !== undefined) newRow[headerIndex.Name] = name;
    if (headerIndex.Manager_ID !== undefined) newRow[headerIndex.Manager_ID] = managerId;
    if (headerIndex.Profile_Pic_Url !== undefined) newRow[headerIndex.Profile_Pic_Url] = profilePicUrl;

    appendRows(sheet, [newRow]);
    invalidateTableCache('Users');

    const createdUser = {};
    headers.forEach((header, index) => {
      createdUser[header] = newRow[index];
    });
    sendManagerAccountCreatedNotification(createdUser);

    return JSON.stringify({ success: true, user: createdUser, created: true });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to add user.'
    });
  }
}

/**
 * API Endpoint: Fetches the full data payload for the frontend to initialize.
 */
function getInitialPayload() {
  const currentUserEmail = normalizeEmail(getCurrentUser());
  const currentUserProfilePhotoUrl = getCurrentUserProfilePhotoUrl();
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  let users = [];

  if (usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const headerIndex = getHeaderIndex(headers);

    if (syncCurrentUserProfilePhoto(usersSheet, data, headerIndex, currentUserEmail, currentUserProfilePhotoUrl)) {
      invalidateTableCache('Users');
    }
    users = getTableData('Users');
  }

  const currentUserExists = users.some((user) => normalizeEmail(user.Email) === currentUserEmail);
  const completionMetadataByTaskId = getTaskCompletionMetadataMap();
  const payload = {
    currentUserEmail: currentUserEmail,
    currentUserExists: currentUserExists,
    requiresAccountSetup: Boolean(currentUserEmail) && !currentUserExists,
    users: users,
    projects: getTableData('Projects'),
    tasks: getTableData('Tasks').map((task) => appendTaskCompletionMetadataToTask(task, completionMetadataByTaskId)),
    assignments: getTableData('Assignments'),
    currentUserSettings: getUserSettingsByEmail(currentUserEmail),
    versionHash: buildGlobalVersionHash(),
    lastUpdated: getStoredDataVersion()
  };

  // Stringifying prevents Apps Script's silent serialization failures
  return JSON.stringify(payload);
}

function getCommentsByTopic(topicId) {
  const normalizedTopicId = topicId ? topicId.toString().trim() : '';
  if (!normalizedTopicId) {
    return JSON.stringify({ success: false, error: 'Topic_ID is required.', comments: [] });
  }

  try {
    const comments = getTableData('Comments')
      .filter((comment) => (comment.Topic_ID || '').toString().trim() === normalizedTopicId)
      .map((comment) => {
        const normalized = {};
        Object.keys(comment).forEach((key) => {
          normalized[key] = normalizeValueForClient(comment[key]);
        });
        return normalized;
      });
    return JSON.stringify({ success: true, topicId: normalizedTopicId, comments });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to load comments.',
      comments: []
    });
  }
}

function getUserById(userId) {
  const normalizedUserId = userId ? userId.toString().trim() : '';
  if (!normalizedUserId) return null;

  const users = getTableData('Users');
  return users.find((user) => (user.User_ID || '').toString().trim() === normalizedUserId) || null;
}

function getProjectById(projectId) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) return null;

  const projects = getTableData('Projects');
  return projects.find((project) => (project.Project_ID || '').toString().trim() === normalizedProjectId) || null;
}

function getTaskById(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) return null;

  const tasks = getTableData('Tasks');
  return tasks.find((task) => (task.Task_ID || '').toString().trim() === normalizedTaskId) || null;
}

function getTaskNotificationDetails(task, projectsById) {
  const taskTitle = (task && task.Task_Title ? task.Task_Title : '').toString().trim() || 'Untitled task';
  const projectId = (task && task.Project_ID ? task.Project_ID : '').toString().trim();
  const project = projectsById
    ? projectsById[projectId]
    : getProjectById(projectId);
  const projectTitle = (project && project.Project_Title ? project.Project_Title : '').toString().trim() || 'Unassigned project';

  return {
    taskTitle,
    projectTitle
  };
}

function formatTaskNotificationBody(details) {
  const taskTitle = details && details.taskTitle ? details.taskTitle : 'Untitled task';
  const projectTitle = details && details.projectTitle ? details.projectTitle : 'Unassigned project';
  return `Task: ${taskTitle}\nProject: ${projectTitle}`;
}

function ensureTaskExists(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    throw new Error('Topic_ID (task ID) is required.');
  }

  const tasks = getTableData('Tasks');
  const taskExists = tasks.some((task) => (task.Task_ID || '').toString().trim() === normalizedTaskId);
  if (!taskExists) {
    throw new Error(`Task ${normalizedTaskId} does not exist.`);
  }

  return normalizedTaskId;
}

function extractMentionedUsers(content, users) {
  if (!content) return [];
  const mentionPattern = /@([a-zA-Z0-9._-]+)/g;
  const matches = content.matchAll(mentionPattern);
  const mentions = new Set();
  for (const match of matches) {
    if (match && match[1]) {
      mentions.add(match[1].toLowerCase());
    }
  }
  if (mentions.size === 0) return [];

  const toHandle = (user) => {
    const emailLocalPart = normalizeEmail(user && user.Email).split('@')[0];
    const fallbackHandle = (user && user.Name ? user.Name : '')
      .toString()
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '.')
      .replace(/[^a-z0-9._-]/g, '');
    return emailLocalPart || fallbackHandle;
  };

  return users.filter((user) => {
    const handle = toHandle(user);
    return handle && mentions.has(handle);
  });
}

function sendMentionNotifications(comment, topicId, commenter, mentionedUsers) {
  if (!mentionedUsers || mentionedUsers.length === 0) return;

  const task = getTaskById(topicId);
  const details = getTaskNotificationDetails(task);
  const commenterName = commenter && commenter.Name ? commenter.Name : 'A teammate';
  const subject = `You were mentioned on ${details.taskTitle}`;
  const body = `${commenterName} mentioned you in a comment.

${formatTaskNotificationBody(details)}

Comment:
${comment}
`;

  const sentEmails = new Set();
  mentionedUsers.forEach((user) => {
    const email = normalizeEmail(user.Email);
    if (!email || sentEmails.has(email)) return;
    if (!isNotificationEnabledForUser(user, 'commentsAndMentions')) return;
    sentEmails.add(email);
    safeSendEmail(email, subject, body);
  });
}

function sendTaskAssignmentNotifications(task, assigneeIds, assignedByUserId) {
  if (!task || !Array.isArray(assigneeIds) || assigneeIds.length === 0) return;

  const usersById = getUsersById();
  const assignedBy = usersById[(assignedByUserId || '').toString().trim()];
  const assignedByName = assignedBy && assignedBy.Name ? assignedBy.Name : 'A teammate';
  const details = getTaskNotificationDetails(task);
  const subject = `New task assignment: ${details.taskTitle}`;

  const sentEmails = new Set();
  assigneeIds.forEach((assigneeId) => {
    const user = usersById[(assigneeId || '').toString().trim()];
    const email = normalizeEmail(user && user.Email);
    if (!email || sentEmails.has(email)) return;
    if (!isNotificationEnabledForUser(user, 'taskAssignments')) return;

    const recipientName = user && user.Name ? user.Name : 'there';
    const body = `Hi ${recipientName},

${assignedByName} assigned you to a task.

${formatTaskNotificationBody(details)}
`;

    sentEmails.add(email);
    safeSendEmail(email, subject, body);
  });
}

function sendManagerTaskCompletedNotifications(task, assigneeIds, completedByUserId) {
  if (!task || !Array.isArray(assigneeIds) || assigneeIds.length === 0) return;

  const users = getTableData('Users');
  const usersById = users.reduce((acc, user) => {
    const userId = (user.User_ID || '').toString().trim();
    if (userId) acc[userId] = user;
    return acc;
  }, {});

  const managerIds = new Set();
  assigneeIds.forEach((assigneeId) => {
    const worker = usersById[(assigneeId || '').toString().trim()];
    const managerId = (worker && worker.Manager_ID ? worker.Manager_ID : '').toString().trim();
    if (managerId) managerIds.add(managerId);
  });
  if (managerIds.size === 0) return;

  const details = getTaskNotificationDetails(task);
  const completedBy = usersById[(completedByUserId || '').toString().trim()];
  const completedByName = completedBy && completedBy.Name ? completedBy.Name : 'A user';
  const subject = `Task completed: ${details.taskTitle}`;
  const body = `${completedByName} marked a task as completed.

${formatTaskNotificationBody(details)}
`;

  const sentEmails = new Set();
  managerIds.forEach((managerId) => {
    const manager = usersById[managerId];
    const email = normalizeEmail(manager && manager.Email);
    if (!email || sentEmails.has(email)) return;
    sentEmails.add(email);
    safeSendEmail(email, subject, body);
  });
}

function sendManagerAccountCreatedNotification(createdUser) {
  if (!createdUser) return;
  const managerId = (createdUser.Manager_ID || '').toString().trim();
  if (!managerId) return;

  const manager = getUserById(managerId);
  const managerEmail = normalizeEmail(manager && manager.Email);
  if (!managerEmail) return;

  const userId = (createdUser.User_ID || '').toString().trim();
  const userName = (createdUser.Name || '').toString().trim() || userId || 'A user';
  const userEmail = normalizeEmail(createdUser.Email);
  const subject = `New direct report created: ${userName}`;
  const body = `${userName}${userEmail ? ` (${userEmail})` : ''} created an account and listed you as their manager.\n\nUser ID: ${userId || 'N/A'}\n`;
  safeSendEmail(managerEmail, subject, body);
}

function addComment(topicId, commentInput) {
  try {
    const normalizedTopicId = ensureTaskExists(topicId);
    const content = commentInput && commentInput.content ? commentInput.content.toString().trim() : '';
    if (!content) {
      throw new Error('Comment content is required.');
    }

    const currentUserEmail = normalizeEmail(getCurrentUser());
    const currentUserId = getCurrentUserIdByEmail(currentUserEmail);
    if (!currentUserId) {
      throw new Error('Could not determine commenter_ID from current user email.');
    }

    const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comments');
    if (!commentsSheet) {
      throw new Error('Comments sheet was not found.');
    }

    const headers = commentsSheet.getRange(1, 1, 1, commentsSheet.getLastColumn()).getValues()[0];
    const headerIndex = getHeaderIndex(headers);
    const now = new Date();
    const newRow = new Array(headers.length).fill('');

    if (headerIndex.Comment_ID !== undefined) newRow[headerIndex.Comment_ID] = generateNextId('Comments', 'C');
    if (headerIndex.Topic_ID !== undefined) newRow[headerIndex.Topic_ID] = normalizedTopicId;
    if (headerIndex.Commenter_ID !== undefined) newRow[headerIndex.Commenter_ID] = currentUserId;
    if (headerIndex.Timestamp !== undefined) newRow[headerIndex.Timestamp] = now;
    if (headerIndex.Content !== undefined) newRow[headerIndex.Content] = content;

    appendRows(commentsSheet, [newRow]);
    invalidateTableCache('Comments');

    const createdComment = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdComment[header] = value instanceof Date ? value.toISOString() : value;
    });

    const users = getTableData('Users');
    const commenter = users.find((user) => (user.User_ID || '').toString().trim() === currentUserId) || null;
    const mentionedUsers = extractMentionedUsers(content, users).filter(
      (user) => (user.User_ID || '').toString().trim() !== currentUserId
    );
    sendMentionNotifications(content, normalizedTopicId, commenter, mentionedUsers);

    return JSON.stringify({ success: true, comment: createdComment });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to add comment.'
    });
  }
}

function deleteComment(commentId) {
  const normalizedCommentId = commentId ? commentId.toString().trim() : '';
  if (!normalizedCommentId) {
    return JSON.stringify({ success: false, error: 'Comment ID is required.' });
  }

  const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comments');
  if (!commentsSheet) {
    return JSON.stringify({ success: false, error: 'Comments sheet was not found.' });
  }

  try {
    const data = commentsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Comments sheet has no data rows.');
    }

    const headers = data[0];
    const commentIdColumnIndex = headers.indexOf('Comment_ID');
    if (commentIdColumnIndex === -1) {
      throw new Error('Comments sheet is missing Comment_ID column.');
    }

    const rowIndex = data.findIndex((row, index) => index > 0 && row[commentIdColumnIndex] === normalizedCommentId);
    if (rowIndex < 0) {
      throw new Error('Comment not found.');
    }

    deleteRowsBySheetIndexes(commentsSheet, [rowIndex + 1]);
    invalidateTableCache('Comments');
    return JSON.stringify({ success: true, commentId: normalizedCommentId });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to delete comment.'
    });
  }
}

function resolveComment(commentId) {
  const normalizedCommentId = commentId ? commentId.toString().trim() : '';
  if (!normalizedCommentId) {
    return JSON.stringify({ success: false, error: 'Comment ID is required.' });
  }

  const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comments');
  if (!commentsSheet) {
    return JSON.stringify({ success: false, error: 'Comments sheet was not found.' });
  }

  try {
    const data = commentsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Comments sheet has no data rows.');
    }

    const headers = data[0];
    const commentIdColumnIndex = headers.indexOf('Comment_ID');
    const topicIdColumnIndex = headers.indexOf('Topic_ID');
    const commenterIdColumnIndex = headers.indexOf('Commenter_ID');
    const contentColumnIndex = headers.indexOf('Content');

    if (commentIdColumnIndex === -1) throw new Error('Comments sheet is missing Comment_ID column.');
    if (commenterIdColumnIndex === -1) throw new Error('Comments sheet is missing Commenter_ID column.');

    const rowIndex = data.findIndex((row, index) => index > 0 && row[commentIdColumnIndex] === normalizedCommentId);
    if (rowIndex < 0) {
      throw new Error('Comment not found.');
    }

    const row = data[rowIndex];
    const commenterId = commenterIdColumnIndex > -1 ? row[commenterIdColumnIndex] : '';
    const commenter = getUserById(commenterId);
    const topicId = topicIdColumnIndex > -1 ? row[topicIdColumnIndex] : '';
    const content = contentColumnIndex > -1 ? row[contentColumnIndex] : '';

    deleteRowsBySheetIndexes(commentsSheet, [rowIndex + 1]);
    invalidateTableCache('Comments');

    const commenterEmail = commenter ? normalizeEmail(commenter.Email) : '';
    if (commenterEmail && isNotificationEnabledForUser(commenter, 'commentsAndMentions')) {
      const resolverEmail = normalizeEmail(getCurrentUser());
      const resolverUserId = getCurrentUserIdByEmail(resolverEmail);
      const resolver = getUserById(resolverUserId);
      const resolverName = resolver && resolver.Name ? resolver.Name : 'A teammate';
      const details = getTaskNotificationDetails(getTaskById(topicId));
      const subject = `Your comment was resolved on ${details.taskTitle}`;
      const body = `${resolverName} resolved your comment.

${formatTaskNotificationBody(details)}

Resolved comment:
${content}
`;
      safeSendEmail(commenterEmail, subject, body);
    }

    return JSON.stringify({ success: true, commentId: normalizedCommentId, resolved: true });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to resolve comment.'
    });
  }
}

function hasDueDatePassed(dueDateValue) {
  if (!dueDateValue) return false;
  const dueDate = new Date(dueDateValue);
  if (dueDate.toString() === 'Invalid Date') return false;

  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const normalizedDueDate = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());

  return normalizedDueDate.getTime() < startOfToday.getTime();
}


function getDaysUntilDate(dateValue) {
  if (!dateValue) return null;
  const dueDate = parseDateInput(dateValue);
  if (!dueDate || dueDate.toString() === 'Invalid Date') return null;

  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const normalizedDueDate = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
  return Math.round((normalizedDueDate.getTime() - startOfToday.getTime()) / (24 * 60 * 60 * 1000));
}

function getAssignmentsByAssignmentId() {
  return getTableData('Assignments').reduce((acc, assignment) => {
    const assignmentId = (assignment.Assignment_ID || '').toString().trim();
    const assigneeId = (assignment.Assignee_ID || '').toString().trim();
    if (!assignmentId || !assigneeId) return acc;
    if (!acc[assignmentId]) acc[assignmentId] = [];
    acc[assignmentId].push(assigneeId);
    return acc;
  }, {});
}

function getUsersById() {
  return getTableData('Users').reduce((acc, user) => {
    const userId = (user.User_ID || '').toString().trim();
    if (userId) acc[userId] = user;
    return acc;
  }, {});
}

function sendDueDateReminderNotifications() {
  const tasks = getTableData('Tasks');
  const assignmentsByTaskId = getAssignmentsByAssignmentId();
  const usersById = getUsersById();
  const scriptProperties = PropertiesService.getScriptProperties();
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let sentCount = 0;

  tasks.forEach((task) => {
    const status = (task.Status || '').toString().trim();
    if (status === 'Completed') return;

    const daysUntilDue = getDaysUntilDate(task.Due_Date);
    if (daysUntilDue === null || daysUntilDue < 0 || daysUntilDue > 1) return;

    const taskId = (task.Task_ID || '').toString().trim();
    const details = getTaskNotificationDetails(task);
    const dueLabel = daysUntilDue === 0 ? 'today' : 'tomorrow';
    const subject = `Task due ${dueLabel}: ${details.taskTitle}`;
    const body = `A task is due ${dueLabel}.

${formatTaskNotificationBody(details)}
`;
    const sentEmails = new Set();

    (assignmentsByTaskId[taskId] || []).forEach((assigneeId) => {
      const user = usersById[assigneeId];
      const email = normalizeEmail(user && user.Email);
      if (!email || sentEmails.has(email)) return;
      if (!isNotificationEnabledForUser(user, 'dueDateReminders')) return;

      const reminderKey = `soar_due_date_reminder:${todayKey}:${taskId}:${email}`;
      if (scriptProperties.getProperty(reminderKey)) return;

      sentEmails.add(email);
      if (safeSendEmail(email, subject, body)) {
        scriptProperties.setProperty(reminderKey, new Date().toISOString());
        sentCount += 1;
      }
    });
  });

  return sentCount;
}

function sendWeeklyDigestNotifications() {
  const tasks = getTableData('Tasks');
  const assignmentsByTaskId = getAssignmentsByAssignmentId();
  const usersById = getUsersById();
  const tasksByUserId = {};

  tasks.forEach((task) => {
    const status = (task.Status || '').toString().trim();
    if (status === 'Completed') return;

    const taskId = (task.Task_ID || '').toString().trim();
    (assignmentsByTaskId[taskId] || []).forEach((assigneeId) => {
      if (!tasksByUserId[assigneeId]) tasksByUserId[assigneeId] = [];
      tasksByUserId[assigneeId].push(task);
    });
  });

  const digestDateKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const scriptProperties = PropertiesService.getScriptProperties();
  let sentCount = 0;

  Object.keys(tasksByUserId).forEach((userId) => {
    const user = usersById[userId];
    const email = normalizeEmail(user && user.Email);
    if (!email || !isNotificationEnabledForUser(user, 'weeklyDigest')) return;

    const digestKey = `soar_weekly_digest:${digestDateKey}:${email}`;
    if (scriptProperties.getProperty(digestKey)) return;

    const taskLines = tasksByUserId[userId]
      .sort((left, right) => {
        const leftDays = getDaysUntilDate(left.Due_Date);
        const rightDays = getDaysUntilDate(right.Due_Date);
        return (leftDays === null ? 9999 : leftDays) - (rightDays === null ? 9999 : rightDays);
      })
      .map((task) => {
        const details = getTaskNotificationDetails(task);
        const dueDate = task.Due_Date ? serializeDateOnlyForClient(parseDateInput(task.Due_Date)) : 'No due date';
        return `- ${details.taskTitle} (${details.projectTitle}), due: ${dueDate}`;
      });

    const subject = 'Your weekly Soar task digest';
    const body = 'Here are your open Soar tasks for this week:\n\n' + taskLines.join('\n') + '\n';
    if (safeSendEmail(email, subject, body)) {
      scriptProperties.setProperty(digestKey, new Date().toISOString());
      sentCount += 1;
    }
  });

  return sentCount;
}

function deleteTaskAndAssignments(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    throw new Error('Task ID is required.');
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = spreadsheet.getSheetByName('Tasks');
  const assignmentsSheet = spreadsheet.getSheetByName('Assignments');

  if (!tasksSheet) throw new Error('Tasks sheet was not found.');
  if (!assignmentsSheet) throw new Error('Assignments sheet was not found.');

  const taskData = tasksSheet.getDataRange().getValues();
  if (taskData.length <= 1) {
    throw new Error('Tasks sheet has no data rows.');
  }

  const headers = taskData[0];
  const taskIdColumnIndex = headers.indexOf('Task_ID');
  if (taskIdColumnIndex === -1) {
    throw new Error('Tasks sheet is missing Task_ID column.');
  }

  const taskRowIndex = taskData.findIndex((row, index) => index > 0 && row[taskIdColumnIndex] === normalizedTaskId);
  if (taskRowIndex < 0) {
    throw new Error('Task not found.');
  }

  deleteRowsBySheetIndexes(tasksSheet, [taskRowIndex + 1]);
  invalidateTableCache('Tasks');

  const assignmentsData = assignmentsSheet.getDataRange().getValues();
  const assignmentRowsToDelete = [];
  for (let i = assignmentsData.length - 1; i >= 1; i--) {
    if (assignmentsData[i][0] === normalizedTaskId) {
      assignmentRowsToDelete.push(i + 1);
    }
  }
  deleteRowsBySheetIndexes(assignmentsSheet, assignmentRowsToDelete);
  invalidateTableCache('Assignments');

  removeTaskCompletionMetadata(normalizedTaskId);

  return normalizedTaskId;
}

function purgeCompletedTasksPastDue() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tasksSheet = spreadsheet.getSheetByName('Tasks');
  const assignmentsSheet = spreadsheet.getSheetByName('Assignments');

  if (!tasksSheet || !assignmentsSheet) return [];

  const taskData = tasksSheet.getDataRange().getValues();
  if (taskData.length <= 1) return [];

  const headers = taskData[0];
  const taskIdColumnIndex = headers.indexOf('Task_ID');
  const statusColumnIndex = headers.indexOf('Status');
  const dueDateColumnIndex = headers.indexOf('Due_Date');
  if (taskIdColumnIndex === -1 || statusColumnIndex === -1 || dueDateColumnIndex === -1) return [];

  const deletedTaskIds = [];
  const taskRowsToDelete = [];
  for (let i = taskData.length - 1; i >= 1; i--) {
    const status = (taskData[i][statusColumnIndex] || '').toString().trim();
    const dueDateValue = taskData[i][dueDateColumnIndex];
    if (status === 'Completed' && hasDueDatePassed(dueDateValue)) {
      deletedTaskIds.push(taskData[i][taskIdColumnIndex]);
      taskRowsToDelete.push(i + 1);
    }
  }
  deleteRowsBySheetIndexes(tasksSheet, taskRowsToDelete);
  invalidateTableCache('Tasks');

  if (deletedTaskIds.length === 0) return [];

  const assignmentIdSet = new Set(deletedTaskIds.map((id) => id.toString().trim()));
  const assignmentsData = assignmentsSheet.getDataRange().getValues();
  const assignmentRowsToDelete = [];
  for (let i = assignmentsData.length - 1; i >= 1; i--) {
    if (assignmentIdSet.has((assignmentsData[i][0] || '').toString().trim())) {
      assignmentRowsToDelete.push(i + 1);
    }
  }
  deleteRowsBySheetIndexes(assignmentsSheet, assignmentRowsToDelete);
  invalidateTableCache('Assignments');

  deletedTaskIds.forEach((taskId) => removeTaskCompletionMetadata(taskId));

  return deletedTaskIds;
}

/**
 * API Endpoint: Updates the current user's profile data.
 */
function updateCurrentUserProfile(profileInput) {
  const activeEmail = normalizeEmail(getCurrentUser());
  const normalizedName = profileInput && profileInput.name ? profileInput.name.toString().trim() : '';
  const normalizedEmail = normalizeEmail(profileInput ? profileInput.email : '');

  if (!normalizedName) {
    return JSON.stringify({ success: false, error: 'Name is required.' });
  }

  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!usersSheet) {
    return JSON.stringify({ success: false, error: 'Users sheet was not found.' });
  }

  try {
    const data = usersSheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Users sheet has no data rows.');
    }

    const headers = data[0];
    const headerIndex = getNormalizedHeaderIndex(headers);

    if (headerIndex.email === undefined) throw new Error('Users sheet is missing Email column.');
    if (headerIndex.name === undefined) throw new Error('Users sheet is missing Name column.');

    const currentUserRowIndex = data.findIndex(
      (row, index) => index > 0 && getUserEmailFromRow(row, headerIndex) === activeEmail
    );
    if (currentUserRowIndex < 0) {
      throw new Error('Current user record was not found.');
    }

    const updatedRow = data[currentUserRowIndex].slice();
    updatedRow[headerIndex.name] = normalizedName;

    if (headerIndex.profile_pic_url !== undefined) {
      const latestProfilePicUrl = getCurrentUserProfilePhotoUrl();
      if (latestProfilePicUrl) {
        updatedRow[headerIndex.profile_pic_url] = latestProfilePicUrl;
      }
    }
    updateRowValues(usersSheet, currentUserRowIndex + 1, updatedRow);
    invalidateTableCache('Users');
    const updatedUser = {};
    headers.forEach((header, index) => {
      const value = updatedRow[index];
      updatedUser[header] = value instanceof Date ? value.toISOString() : value;
    });

    return JSON.stringify({ success: true, user: updatedUser });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update user profile.'
    });
  }
}

/**
 * API Endpoint: Updates a task status.
 */
function updateTaskStatus(taskId, newStatus) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    return JSON.stringify({ success: false, error: 'Task ID is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }

  try {
    const status = normalizeTaskStatus(newStatus);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Tasks sheet has no data rows.');
    }

    const headers = data[0];
    const headerIndex = getHeaderIndex(headers);
    const taskIdColumnIndex = headerIndex.Task_ID;
    const statusColumnIndex = headerIndex.Status;

    if (taskIdColumnIndex === undefined || statusColumnIndex === undefined) {
      throw new Error('Tasks sheet is missing Task_ID or Status column.');
    }

    const taskRowIndex = data.findIndex((row, index) => index > 0 && (row[taskIdColumnIndex] || '').toString().trim() === normalizedTaskId);
    if (taskRowIndex < 0) {
      throw new Error('Task not found.');
    }

    const previousStatus = (data[taskRowIndex][statusColumnIndex] || '').toString().trim();
    const isMarkingCompleted = status === 'Completed' && previousStatus !== 'Completed';
    const isReopeningCompleted = status !== 'Completed' && previousStatus === 'Completed';
    const completedAt = new Date();
    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    const updatedRow = data[taskRowIndex].slice();

    updatedRow[statusColumnIndex] = status;
    if (headerIndex.Completed_By !== undefined) {
      if (isMarkingCompleted) {
        updatedRow[headerIndex.Completed_By] = currentUserId || '';
      } else if (isReopeningCompleted) {
        updatedRow[headerIndex.Completed_By] = '';
      }
    }
    if (headerIndex.Completed_At !== undefined) {
      if (isMarkingCompleted) {
        updatedRow[headerIndex.Completed_At] = completedAt;
      } else if (isReopeningCompleted) {
        updatedRow[headerIndex.Completed_At] = '';
      }
    }

    updateRowValues(sheet, taskRowIndex + 1, updatedRow);
    invalidateTableCache('Tasks');

    if (isMarkingCompleted) {
      upsertTaskCompletionMetadata(normalizedTaskId, currentUserId, completedAt);
    } else if (isReopeningCompleted) {
      removeTaskCompletionMetadata(normalizedTaskId);
    }

    const updatedTask = {};
    headers.forEach((header, index) => {
      const value = updatedRow[index];
      updatedTask[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });
    appendTaskCompletionMetadataToTask(updatedTask);

    if (isMarkingCompleted) {
      const assigneeIds = getTableData('Assignments')
        .filter((assignment) => (assignment.Assignment_ID || '').toString().trim() === normalizedTaskId)
        .map((assignment) => (assignment.Assignee_ID || '').toString().trim())
        .filter(Boolean);
      sendManagerTaskCompletedNotifications(updatedTask, assigneeIds, currentUserId);
    }

    return JSON.stringify({ success: true, task: updatedTask });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update task status.'
    });
  }
}

function normalizePriorityValue(value) {
  if (value === null || value === undefined || value === '') return '';

  const validPriorities = ['High', 'Medium', 'Low'];
  const legacyPriorityMap = {
    5: 'High',
    4: 'High',
    3: 'Medium',
    2: 'Low',
    1: 'Low'
  };
  const stringValue = value.toString().trim();

  if (!stringValue) return '';
  if (validPriorities.includes(stringValue)) return stringValue;

  const numericValue = Number(stringValue);
  if (Number.isInteger(numericValue) && legacyPriorityMap[numericValue]) {
    return legacyPriorityMap[numericValue];
  }

  throw new Error(`Priority must be one of: ${validPriorities.join(', ')}.`);
}

function normalizeStatusValue(value) {
  const validStatuses = ['Not Started', 'In Progress', 'Completed', 'Delayed'];
  const status = value && value.toString().trim() ? value.toString().trim() : 'Not Started';
  if (!validStatuses.includes(status)) {
    throw new Error(`Status must be one of: ${validStatuses.join(', ')}`);
  }
  return status;
}

function getCurrentUserIdByEmail(email) {
  if (!email) return '';

  const users = getTableData('Users');
  const normalizedEmail = email.toString().trim().toLowerCase();
  const user = users.find((u) => (u.Email || '').toString().trim().toLowerCase() === normalizedEmail);

  return user && user.User_ID ? user.User_ID : '';
}


function getAssignableUserIdsForUser(actorUserId) {
  const normalizedActorId = (actorUserId || '').toString().trim();
  if (!normalizedActorId) return new Set();

  const users = getTableData('Users');
  if (!Array.isArray(users) || users.length === 0) {
    return new Set([normalizedActorId]);
  }

  const directReportsByManager = users.reduce((acc, user) => {
    const managerId = (user.Manager_ID || '').toString().trim();
    const userId = (user.User_ID || '').toString().trim();
    if (!managerId || !userId) return acc;
    if (!acc[managerId]) acc[managerId] = [];
    acc[managerId].push(userId);
    return acc;
  }, {});

  const assignableIds = new Set([normalizedActorId]);
  const queue = [normalizedActorId];
  while (queue.length > 0) {
    const managerId = queue.shift();
    const reports = directReportsByManager[managerId] || [];
    reports.forEach((reportId) => {
      if (!assignableIds.has(reportId)) {
        assignableIds.add(reportId);
        queue.push(reportId);
      }
    });
  }

  return assignableIds;
}

function getValidAssigneeIds(assigneeIds) {
  if (!Array.isArray(assigneeIds)) return [];

  const uniqueAssigneeIds = [...new Set(assigneeIds.map(id => (id || '').toString().trim()).filter(Boolean))];
  if (uniqueAssigneeIds.length === 0) return [];

  const users = getTableData('Users');
  const validUserIds = new Set(users.map(user => user.User_ID));
  const invalidAssignees = uniqueAssigneeIds.filter(id => !validUserIds.has(id));

  if (invalidAssignees.length > 0) {
    throw new Error(`Invalid assignee ID(s): ${invalidAssignees.join(', ')}`);
  }

  return uniqueAssigneeIds;
}

function ensureTaskHasAssignees(assigneeIds) {
  if (!Array.isArray(assigneeIds) || assigneeIds.length === 0) {
    throw new Error('At least one assignee is required for every task.');
  }
}

function validateAssigneePermissions(assigneeIds, actorUserId, grandfatheredAssigneeIds) {
  const normalizedActorId = (actorUserId || '').toString().trim();
  if (!normalizedActorId) {
    throw new Error('Could not determine current user for assignment permissions.');
  }

  const assignableIds = getAssignableUserIdsForUser(normalizedActorId);
  const grandfatheredIdSet = new Set(
    Array.isArray(grandfatheredAssigneeIds)
      ? grandfatheredAssigneeIds.map((assigneeId) => (assigneeId || '').toString().trim()).filter(Boolean)
      : []
  );
  const invalidAssigneeIds = assigneeIds.filter(
    (assigneeId) => !assignableIds.has(assigneeId) && !grandfatheredIdSet.has(assigneeId)
  );
  if (invalidAssigneeIds.length > 0) {
    throw new Error('You can only assign tasks to yourself or your reporting chain.');
  }
}

/**
 * API Endpoint: Creates a project.
 */
function createProject(projectInput) {
  const projectTitle = projectInput && projectInput.projectTitle ? projectInput.projectTitle.toString().trim() : '';
  if (!projectTitle) {
    return JSON.stringify({ success: false, error: 'Project title is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Projects sheet was not found.' });
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    const newRow = new Array(headers.length).fill('');
    const now = new Date();
    const status = normalizeStatusValue(projectInput ? projectInput.status : '');
    const description = projectInput && projectInput.description ? projectInput.description.toString().trim() : '';
    const parsedDueDate = parseDateInput(projectInput ? projectInput.dueDate : '');
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const creatorId = getCurrentUserIdByEmail(getCurrentUser());

    if (!creatorId) {
      throw new Error('Could not determine Creator_ID from current user email.');
    }

    if (headerIndex.Project_ID !== undefined) newRow[headerIndex.Project_ID] = generateNextId('Projects', 'P');
    if (headerIndex.Project_Title !== undefined) newRow[headerIndex.Project_Title] = projectTitle;
    if (headerIndex.Description !== undefined) newRow[headerIndex.Description] = description;
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = status;
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = now;
    if (headerIndex.Due_Date !== undefined) newRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;

    appendRows(sheet, [newRow]);
    invalidateTableCache('Projects');

    const projectId = headerIndex.Project_ID !== undefined ? newRow[headerIndex.Project_ID] : '';
    if (projectId) {
      const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
      if (!assignmentsSheet) {
        throw new Error('Assignments sheet was not found.');
      }
      appendRows(assignmentsSheet, [[projectId, creatorId]]);
      invalidateTableCache('Assignments');
    }

    const createdProject = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdProject[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });

    const createdAssignment = projectId
      ? { Assignment_ID: projectId, Assignee_ID: creatorId }
      : null;

    return JSON.stringify({ success: true, project: createdProject, assignment: createdAssignment });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to create project.'
    });
  }
}

/**
 * API Endpoint: Creates a task for a project.
 */
function createTask(projectId, taskInput) {
  if (!projectId) {
    return JSON.stringify({ success: false, error: 'Project ID is required.' });
  }

  const taskTitle = taskInput && taskInput.taskTitle ? taskInput.taskTitle.toString().trim() : '';
  if (!taskTitle) {
    return JSON.stringify({ success: false, error: 'Task title is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    const newRow = new Array(headers.length).fill('');
    const now = new Date();
    const parsedDueDate = parseDateInput(taskInput ? taskInput.dueDate : '');
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const priority = normalizePriorityValue(taskInput ? taskInput.priority : '');
    const description = taskInput && taskInput.description ? taskInput.description.toString().trim() : '';
    const assigneeIds = getValidAssigneeIds(taskInput ? taskInput.assigneeIds : []);
    ensureTaskHasAssignees(assigneeIds);
    const normalizedProjectId = ensureProjectExists(projectId);
    const creatorId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));

    if (!creatorId) {
      throw new Error('Could not determine Creator_ID from current user email.');
    }

    validateAssigneePermissions(assigneeIds, creatorId);

    if (headerIndex.Task_ID !== undefined) newRow[headerIndex.Task_ID] = generateNextId('Tasks', 'T');
    if (headerIndex.Project_ID !== undefined) newRow[headerIndex.Project_ID] = normalizedProjectId;
    if (headerIndex.Task_Title !== undefined) newRow[headerIndex.Task_Title] = taskTitle;
    if (headerIndex.Due_Date !== undefined) newRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = 'Not Started';
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = now;
    if (headerIndex.Description !== undefined) newRow[headerIndex.Description] = description;
    if (headerIndex.Priority !== undefined) newRow[headerIndex.Priority] = priority;
    if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;

    appendRows(sheet, [newRow]);
    invalidateTableCache('Tasks');

    let createdAssignments = [];
    if (assigneeIds.length > 0 && headerIndex.Task_ID !== undefined) {
      const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
      if (!assignmentsSheet) {
        throw new Error('Assignments sheet was not found.');
      }

      createdAssignments = assigneeIds.map(assigneeId => ({
        Assignment_ID: newRow[headerIndex.Task_ID],
        Assignee_ID: assigneeId
      }));

      const assignmentRows = createdAssignments.map(assignment => [assignment.Assignment_ID, assignment.Assignee_ID]);
      appendRows(assignmentsSheet, assignmentRows);
      invalidateTableCache('Assignments');
    }

    const createdTask = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdTask[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });
    sendTaskAssignmentNotifications(createdTask, assigneeIds, creatorId);

    return JSON.stringify({ success: true, task: createdTask, assignments: createdAssignments });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to create task.'
    });
  }
}

function normalizeTaskStatus(status) {
  const allowedStatuses = ['Not Started', 'In Progress', 'Completed', 'Delayed'];
  const normalizedStatus = status ? status.toString().trim() : '';

  if (!normalizedStatus) {
    throw new Error('Status is required.');
  }

  if (allowedStatuses.indexOf(normalizedStatus) === -1) {
    throw new Error(`Status must be one of: ${allowedStatuses.join(', ')}.`);
  }

  return normalizedStatus;
}

function ensureProjectExists(projectId) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) {
    throw new Error('Project ID is required.');
  }

  const projectIds = new Set(getTableData('Projects').map(project => project.Project_ID));
  if (!projectIds.has(normalizedProjectId)) {
    throw new Error(`Project ID ${normalizedProjectId} does not exist.`);
  }

  return normalizedProjectId;
}

/**
 * API Endpoint: Updates a task and its assignees.
 */
function updateTask(taskId, taskInput) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    return JSON.stringify({ success: false, error: 'Task ID is required.' });
  }

  const taskTitle = taskInput && taskInput.taskTitle ? taskInput.taskTitle.toString().trim() : '';
  if (!taskTitle) {
    return JSON.stringify({ success: false, error: 'Task title is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }

  try {
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      throw new Error('Tasks sheet has no data rows.');
    }

    const headers = data[0];
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    const taskIdColumnIndex = headerIndex.Task_ID;
    if (taskIdColumnIndex === undefined) {
      throw new Error('Tasks sheet is missing Task_ID column.');
    }

    const taskRowIndex = data.findIndex((row, index) => index > 0 && row[taskIdColumnIndex] === normalizedTaskId);
    if (taskRowIndex < 0) {
      throw new Error('Task not found.');
    }
    const previousStatus = (data[taskRowIndex][headerIndex.Status] || '').toString().trim();

    const parsedDueDate = parseDateInput(taskInput ? taskInput.dueDate : '');
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const priority = normalizePriorityValue(taskInput ? taskInput.priority : '');
    const description = taskInput && taskInput.description ? taskInput.description.toString().trim() : '';
    const assigneeIds = getValidAssigneeIds(taskInput ? taskInput.assigneeIds : []);
    ensureTaskHasAssignees(assigneeIds);
    const status = normalizeTaskStatus(taskInput ? taskInput.status : '');
    const projectId = ensureProjectExists(taskInput ? taskInput.projectId : '');
    const editorId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));

    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    const isMarkingCompleted = status === 'Completed' && previousStatus !== 'Completed';
    const isReopeningCompleted = status !== 'Completed' && previousStatus === 'Completed';

    const updatedTaskRow = data[taskRowIndex].slice();
    if (headerIndex.Project_ID !== undefined) updatedTaskRow[headerIndex.Project_ID] = projectId;
    if (headerIndex.Task_Title !== undefined) updatedTaskRow[headerIndex.Task_Title] = taskTitle;
    if (headerIndex.Due_Date !== undefined) updatedTaskRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Status !== undefined) updatedTaskRow[headerIndex.Status] = status;
    if (headerIndex.Description !== undefined) updatedTaskRow[headerIndex.Description] = description;
    if (headerIndex.Priority !== undefined) updatedTaskRow[headerIndex.Priority] = priority;
    const completedAt = new Date();
    if (headerIndex.Completed_By !== undefined) {
      if (isMarkingCompleted) {
        updatedTaskRow[headerIndex.Completed_By] = currentUserId || '';
      } else if (isReopeningCompleted) {
        updatedTaskRow[headerIndex.Completed_By] = '';
      }
    }
    if (headerIndex.Completed_At !== undefined) {
      if (isMarkingCompleted) {
        updatedTaskRow[headerIndex.Completed_At] = completedAt;
      } else if (isReopeningCompleted) {
        updatedTaskRow[headerIndex.Completed_At] = '';
      }
    }
    updateRowValues(sheet, taskRowIndex + 1, updatedTaskRow);
    invalidateTableCache('Tasks');
    if (isMarkingCompleted) {
      upsertTaskCompletionMetadata(normalizedTaskId, currentUserId, completedAt);
    } else if (isReopeningCompleted) {
      removeTaskCompletionMetadata(normalizedTaskId);
    }

    const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
    if (!assignmentsSheet) {
      throw new Error('Assignments sheet was not found.');
    }

    const assignmentsData = assignmentsSheet.getDataRange().getValues();
    const previousAssigneeIds = assignmentsData
      .filter((row, index) => index > 0 && row[0] === normalizedTaskId)
      .map((row) => (row[1] || '').toString().trim())
      .filter(Boolean);
    validateAssigneePermissions(assigneeIds, editorId, previousAssigneeIds);
    const assignmentRowsToDelete = [];
    for (let i = assignmentsData.length - 1; i >= 1; i--) {
      if (assignmentsData[i][0] === normalizedTaskId) {
        assignmentRowsToDelete.push(i + 1);
      }
    }
    deleteRowsBySheetIndexes(assignmentsSheet, assignmentRowsToDelete);
    invalidateTableCache('Assignments');

    let updatedAssignments = [];
    if (assigneeIds.length > 0) {
      updatedAssignments = assigneeIds.map((assigneeId) => ({
        Assignment_ID: normalizedTaskId,
        Assignee_ID: assigneeId
      }));

      const assignmentRows = updatedAssignments.map((assignment) => [assignment.Assignment_ID, assignment.Assignee_ID]);
      appendRows(assignmentsSheet, assignmentRows);
      invalidateTableCache('Assignments');
    }

    const refreshedRow = updatedTaskRow;

    const updatedTask = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      updatedTask[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });
    appendTaskCompletionMetadataToTask(updatedTask);
    const previousAssigneeSet = new Set(previousAssigneeIds);
    const newlyAssignedIds = assigneeIds.filter((assigneeId) => !previousAssigneeSet.has(assigneeId));
    sendTaskAssignmentNotifications(updatedTask, newlyAssignedIds, currentUserId);
    if (isMarkingCompleted) {
      sendManagerTaskCompletedNotifications(updatedTask, assigneeIds, currentUserId);
    }

    return JSON.stringify({ success: true, task: updatedTask, assignments: updatedAssignments });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update task.'
    });
  }
}

/**
 * API Endpoint: Updates a project's details.
 */
function updateProject(projectId, projectInput) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) {
    return JSON.stringify({ success: false, error: 'Project ID is required.' });
  }

  const projectTitle = projectInput && projectInput.projectTitle ? projectInput.projectTitle.toString().trim() : '';
  if (!projectTitle) {
    return JSON.stringify({ success: false, error: 'Project title is required.' });
  }

  const projectsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if (!projectsSheet) {
    return JSON.stringify({ success: false, error: 'Projects sheet was not found.' });
  }

  try {
    const dataRange = projectsSheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      throw new Error('Projects sheet has no data rows.');
    }

    const headers = data[0];
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    const projectIdColumnIndex = headerIndex.Project_ID;
    if (projectIdColumnIndex === undefined) {
      throw new Error('Projects sheet is missing Project_ID column.');
    }

    const projectRowIndex = data.findIndex((row, index) => index > 0 && row[projectIdColumnIndex] === normalizedProjectId);
    if (projectRowIndex < 0) {
      throw new Error('Project not found.');
    }

    const status = normalizeStatusValue(projectInput ? projectInput.status : '');
    const description = projectInput && projectInput.description ? projectInput.description.toString().trim() : '';
    const parsedDueDate = parseDateInput(projectInput ? projectInput.dueDate : '');
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';

    const refreshedRow = data[projectRowIndex].slice();
    if (headerIndex.Project_Title !== undefined) refreshedRow[headerIndex.Project_Title] = projectTitle;
    if (headerIndex.Description !== undefined) refreshedRow[headerIndex.Description] = description;
    if (headerIndex.Status !== undefined) refreshedRow[headerIndex.Status] = status;
    if (headerIndex.Due_Date !== undefined) refreshedRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    updateRowValues(projectsSheet, projectRowIndex + 1, refreshedRow);
    invalidateTableCache('Projects');

    const updatedProject = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      updatedProject[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });

    return JSON.stringify({ success: true, project: updatedProject });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update project.'
    });
  }
}

/**
 * API Endpoint: Deletes a project and cascades deletion to its tasks and assignments.
 */
function deleteProject(projectId) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) {
    return JSON.stringify({ success: false, error: 'Project ID is required.' });
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = spreadsheet.getSheetByName('Projects');
  const tasksSheet = spreadsheet.getSheetByName('Tasks');
  const assignmentsSheet = spreadsheet.getSheetByName('Assignments');

  if (!projectsSheet) {
    return JSON.stringify({ success: false, error: 'Projects sheet was not found.' });
  }
  if (!tasksSheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }
  if (!assignmentsSheet) {
    return JSON.stringify({ success: false, error: 'Assignments sheet was not found.' });
  }

  try {
    const projectData = projectsSheet.getDataRange().getValues();
    if (projectData.length <= 1) {
      throw new Error('Projects sheet has no data rows.');
    }

    const projectHeaders = projectData[0];
    const projectIdColumnIndex = projectHeaders.indexOf('Project_ID');
    if (projectIdColumnIndex === -1) {
      throw new Error('Projects sheet is missing Project_ID column.');
    }

    const projectRowIndex = projectData.findIndex(
      (row, index) => index > 0 && row[projectIdColumnIndex] === normalizedProjectId
    );
    if (projectRowIndex < 0) {
      throw new Error('Project not found.');
    }

    const tasksData = tasksSheet.getDataRange().getValues();
    const taskHeaders = tasksData[0] || [];
    const taskIdColumnIndex = taskHeaders.indexOf('Task_ID');
    const taskProjectIdColumnIndex = taskHeaders.indexOf('Project_ID');
    if (taskIdColumnIndex === -1 || taskProjectIdColumnIndex === -1) {
      throw new Error('Tasks sheet must contain Task_ID and Project_ID columns.');
    }

    const deletedTaskIds = [];
    const taskRowsToDelete = [];
    for (let i = tasksData.length - 1; i >= 1; i--) {
      if (tasksData[i][taskProjectIdColumnIndex] === normalizedProjectId) {
        deletedTaskIds.push(tasksData[i][taskIdColumnIndex]);
        taskRowsToDelete.push(i + 1);
      }
    }
    deleteRowsBySheetIndexes(tasksSheet, taskRowsToDelete);
    invalidateTableCache('Tasks');

    if (deletedTaskIds.length > 0) {
      const deletedTaskIdSet = new Set(deletedTaskIds);
      const assignmentsData = assignmentsSheet.getDataRange().getValues();
      const assignmentRowsToDelete = [];
      for (let i = assignmentsData.length - 1; i >= 1; i--) {
        if (deletedTaskIdSet.has(assignmentsData[i][0])) {
          assignmentRowsToDelete.push(i + 1);
        }
      }
      deleteRowsBySheetIndexes(assignmentsSheet, assignmentRowsToDelete);
      invalidateTableCache('Assignments');
      deletedTaskIds.forEach((taskId) => removeTaskCompletionMetadata(taskId));
    }

    deleteRowsBySheetIndexes(projectsSheet, [projectRowIndex + 1]);
    invalidateTableCache('Projects');

    return JSON.stringify({
      success: true,
      projectId: normalizedProjectId,
      deletedTaskIds: deletedTaskIds
    });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to delete project.'
    });
  }
}

/**
 * API Endpoint: Deletes a task and any of its assignments.
 */
function deleteTask(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    return JSON.stringify({ success: false, error: 'Task ID is required.' });
  }

  try {
    deleteTaskAndAssignments(normalizedTaskId);
    return JSON.stringify({ success: true, taskId: normalizedTaskId });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to delete task.'
    });
  }
}

function completeTask(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    return JSON.stringify({ success: false, error: 'Task ID is required.' });
  }

  const tasksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!tasksSheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }

  try {
    const taskData = tasksSheet.getDataRange().getValues();
    if (taskData.length <= 1) {
      throw new Error('Tasks sheet has no data rows.');
    }

    const headers = taskData[0];
    const taskIdColumnIndex = headers.indexOf('Task_ID');
    const statusColumnIndex = headers.indexOf('Status');
    const completedByColumnIndex = headers.indexOf('Completed_By');
    const completedAtColumnIndex = headers.indexOf('Completed_At');

    if (taskIdColumnIndex === -1) throw new Error('Tasks sheet is missing Task_ID column.');
    if (statusColumnIndex === -1) throw new Error('Tasks sheet is missing Status column.');

    const taskRowIndex = taskData.findIndex((row, index) => index > 0 && row[taskIdColumnIndex] === normalizedTaskId);
    if (taskRowIndex < 0) throw new Error('Task not found.');

    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    const completedAt = new Date();
    const refreshedRow = taskData[taskRowIndex].slice();
    refreshedRow[statusColumnIndex] = 'Completed';
    if (completedByColumnIndex > -1) refreshedRow[completedByColumnIndex] = currentUserId || '';
    if (completedAtColumnIndex > -1) refreshedRow[completedAtColumnIndex] = completedAt;
    updateRowValues(tasksSheet, taskRowIndex + 1, refreshedRow);
    invalidateTableCache('Tasks');
    upsertTaskCompletionMetadata(normalizedTaskId, currentUserId, completedAt);
    const completedTask = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      completedTask[header] = value instanceof Date ? value.toISOString() : value;
    });
    appendTaskCompletionMetadataToTask(completedTask);
    const assignments = getTableData('Assignments');
    const assigneeIds = assignments
      .filter((assignment) => (assignment.Assignment_ID || '').toString().trim() === normalizedTaskId)
      .map((assignment) => (assignment.Assignee_ID || '').toString().trim())
      .filter(Boolean);
    sendManagerTaskCompletedNotifications(completedTask, assigneeIds, currentUserId);

    return JSON.stringify({ success: true, task: completedTask });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to complete task.'
    });
  }
}
