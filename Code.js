/**
 * Main entry point for the web app.
 * Serves the Index.html file.
 * is version 2 even getting there
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

/**
 * Core DB Function: Reads a sheet and returns an array of JSON objects.
 */
function getTableData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Empty or headers only
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * Generates the zero-padded ID defined in the README.
 */
function generateNextId(sheetName, prefix) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return `${prefix}-00000001`; // First ID
  }
  
  const existingIds = data
    .slice(1)
    .map(row => (row[0] || '').toString())
    .filter(id => id.startsWith(`${prefix}-`));

  if (existingIds.length === 0) {
    return `${prefix}-00000001`;
  }

  const maxNumber = existingIds.reduce((max, id) => {
    const numericPart = parseInt(id.split('-')[1], 10);
    if (Number.isNaN(numericPart)) return max;
    return Math.max(max, numericPart);
  }, 0);

  const nextNumber = maxNumber + 1;
  
  return `${prefix}-${nextNumber.toString().padStart(8, '0')}`;
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

    sheet.appendRow(newRow);

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
  purgeCompletedTasksPastDue();

  const currentUserEmail = normalizeEmail(getCurrentUser());
  const currentUserProfilePhotoUrl = getCurrentUserProfilePhotoUrl();
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  let users = [];

  if (usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const headerIndex = getHeaderIndex(headers);

    syncCurrentUserProfilePhoto(usersSheet, data, headerIndex, currentUserEmail, currentUserProfilePhotoUrl);
    users = getTableData('Users');
  }

  const currentUserExists = users.some((user) => normalizeEmail(user.Email) === currentUserEmail);
  const payload = {
    currentUserEmail: currentUserEmail,
    currentUserExists: currentUserExists,
    requiresAccountSetup: Boolean(currentUserEmail) && !currentUserExists,
    users: users,
    projects: getTableData('Projects'),
    tasks: getTableData('Tasks'),
    assignments: getTableData('Assignments'),
    comments: getTableData('Comments')
  };
  
  // Stringifying prevents Apps Script's silent serialization failures
  return JSON.stringify(payload); 
}

function getUserById(userId) {
  const normalizedUserId = userId ? userId.toString().trim() : '';
  if (!normalizedUserId) return null;

  const users = getTableData('Users');
  return users.find((user) => (user.User_ID || '').toString().trim() === normalizedUserId) || null;
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

  const topicLabel = topicId || 'task';
  const commenterName = commenter && commenter.Name ? commenter.Name : 'A teammate';
  const commenterEmail = commenter && commenter.Email ? commenter.Email : 'Unknown email';
  const subject = `You were mentioned in a comment on ${topicLabel}`;
  const body = `${commenterName} (${commenterEmail}) mentioned you in a comment on ${topicLabel}.\n\nComment:\n${comment}\n`;

  const sentEmails = new Set();
  mentionedUsers.forEach((user) => {
    const email = normalizeEmail(user.Email);
    if (!email || sentEmails.has(email)) return;
    sentEmails.add(email);
    MailApp.sendEmail(email, subject, body);
  });
}

function sendTaskAssignmentNotifications(task, assigneeIds) {
  if (!task || !Array.isArray(assigneeIds) || assigneeIds.length === 0) return;

  const usersById = getTableData('Users').reduce((acc, user) => {
    const userId = (user.User_ID || '').toString().trim();
    if (userId) acc[userId] = user;
    return acc;
  }, {});

  const taskId = (task.Task_ID || '').toString().trim();
  const taskTitle = (task.Task_Title || '').toString().trim() || taskId || 'a task';
  const projectId = (task.Project_ID || '').toString().trim();
  const subject = `You were assigned to task ${taskId || taskTitle}`;
  const body = `You were assigned to task "${taskTitle}"${projectId ? ` in project ${projectId}` : ''}.\n\nTask ID: ${taskId || 'N/A'}\n`;

  const sentEmails = new Set();
  assigneeIds.forEach((assigneeId) => {
    const user = usersById[(assigneeId || '').toString().trim()];
    const email = normalizeEmail(user && user.Email);
    if (!email || sentEmails.has(email)) return;
    sentEmails.add(email);
    MailApp.sendEmail(email, subject, body);
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

  const taskId = (task.Task_ID || '').toString().trim();
  const taskTitle = (task.Task_Title || '').toString().trim() || taskId || 'a task';
  const completedBy = usersById[(completedByUserId || '').toString().trim()];
  const completedByName = completedBy && completedBy.Name ? completedBy.Name : 'A user';
  const subject = `Worker task completed: ${taskId || taskTitle}`;
  const body = `${completedByName} marked task "${taskTitle}" as completed.\n\nTask ID: ${taskId || 'N/A'}\n`;

  const sentEmails = new Set();
  managerIds.forEach((managerId) => {
    const manager = usersById[managerId];
    const email = normalizeEmail(manager && manager.Email);
    if (!email || sentEmails.has(email)) return;
    sentEmails.add(email);
    MailApp.sendEmail(email, subject, body);
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
  MailApp.sendEmail(managerEmail, subject, body);
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

    commentsSheet.appendRow(newRow);

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

    commentsSheet.deleteRow(rowIndex + 1);
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

    commentsSheet.deleteRow(rowIndex + 1);

    const commenterEmail = commenter ? normalizeEmail(commenter.Email) : '';
    if (commenterEmail) {
      const resolverEmail = normalizeEmail(getCurrentUser());
      const resolverUserId = getCurrentUserIdByEmail(resolverEmail);
      const resolver = getUserById(resolverUserId);
      const resolverName = resolver && resolver.Name ? resolver.Name : 'A teammate';
      const subject = `Your comment ${normalizedCommentId} was resolved`;
      const body = `${resolverName} resolved your comment on ${topicId || 'a task'}.\n\nResolved comment:\n${content}\n`;
      MailApp.sendEmail(commenterEmail, subject, body);
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

  tasksSheet.deleteRow(taskRowIndex + 1);

  const assignmentsData = assignmentsSheet.getDataRange().getValues();
  for (let i = assignmentsData.length - 1; i >= 1; i--) {
    if (assignmentsData[i][0] === normalizedTaskId) {
      assignmentsSheet.deleteRow(i + 1);
    }
  }

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
  for (let i = taskData.length - 1; i >= 1; i--) {
    const status = (taskData[i][statusColumnIndex] || '').toString().trim();
    const dueDateValue = taskData[i][dueDateColumnIndex];
    if (status === 'Completed' && hasDueDatePassed(dueDateValue)) {
      deletedTaskIds.push(taskData[i][taskIdColumnIndex]);
      tasksSheet.deleteRow(i + 1);
    }
  }

  if (deletedTaskIds.length === 0) return [];

  const assignmentIdSet = new Set(deletedTaskIds.map((id) => id.toString().trim()));
  const assignmentsData = assignmentsSheet.getDataRange().getValues();
  for (let i = assignmentsData.length - 1; i >= 1; i--) {
    if (assignmentIdSet.has((assignmentsData[i][0] || '').toString().trim())) {
      assignmentsSheet.deleteRow(i + 1);
    }
  }

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

    const duplicateEmailIndex = data.findIndex(
      (row, index) =>
        index > 0 &&
        index !== currentUserRowIndex &&
        getUserEmailFromRow(row, headerIndex) === normalizedEmail
    );
    if (duplicateEmailIndex > 0) {
      throw new Error('Email already exists for another user.');
    }

    usersSheet.getRange(currentUserRowIndex + 1, headerIndex.name + 1).setValue(normalizedName);
    usersSheet.getRange(currentUserRowIndex + 1, headerIndex.email + 1).setValue(normalizedEmail);
    if (headerIndex.profile_pic_url !== undefined) {
      const latestProfilePicUrl = getCurrentUserProfilePhotoUrl();
      if (latestProfilePicUrl) {
        usersSheet.getRange(currentUserRowIndex + 1, headerIndex.profile_pic_url + 1).setValue(latestProfilePicUrl);
      }
    }

    const updatedRow = usersSheet.getRange(currentUserRowIndex + 1, 1, 1, headers.length).getValues()[0];
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) {
      // Assuming 'Status' is the 5th column (index 4) based on Data Dictionary
      sheet.getRange(i + 1, 5).setValue(newStatus);
      return { success: true, taskId: taskId, newStatus: newStatus };
    }
  }
  return { success: false, error: "Task not found" };
}

function normalizeScaleValue(value, fieldName) {
  if (value === null || value === undefined || value === '') return '';

  const parsedValue = Number(value);
  if (!Number.isInteger(parsedValue) || parsedValue < 1 || parsedValue > 5) {
    throw new Error(`${fieldName} must be an integer between 1 and 5.`);
  }
  return parsedValue;
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
    const parsedDueDate = projectInput && projectInput.dueDate ? new Date(projectInput.dueDate) : '';
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

    sheet.appendRow(newRow);

    const createdProject = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdProject[header] = value instanceof Date ? value.toISOString() : value;
    });

    return JSON.stringify({ success: true, project: createdProject });
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
    const parsedDueDate = taskInput && taskInput.dueDate ? new Date(taskInput.dueDate) : '';
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const complexity = normalizeScaleValue(taskInput ? taskInput.complexity : '', 'Complexity');
    const priority = normalizeScaleValue(taskInput ? taskInput.priority : '', 'Priority');
    const description = taskInput && taskInput.description ? taskInput.description.toString().trim() : '';
    const assigneeIds = getValidAssigneeIds(taskInput ? taskInput.assigneeIds : []);
    const normalizedProjectId = ensureProjectExists(projectId);

    if (headerIndex.Task_ID !== undefined) newRow[headerIndex.Task_ID] = generateNextId('Tasks', 'T');
    if (headerIndex.Project_ID !== undefined) newRow[headerIndex.Project_ID] = normalizedProjectId;
    if (headerIndex.Task_Title !== undefined) newRow[headerIndex.Task_Title] = taskTitle;
    if (headerIndex.Due_Date !== undefined) newRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = 'Not Started';
    if (headerIndex.Complexity !== undefined) newRow[headerIndex.Complexity] = complexity;
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = now;
    if (headerIndex.Description !== undefined) newRow[headerIndex.Description] = description;
    if (headerIndex.Priority !== undefined) newRow[headerIndex.Priority] = priority;
    if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;

    sheet.appendRow(newRow);

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
      assignmentsSheet.getRange(assignmentsSheet.getLastRow() + 1, 1, assignmentRows.length, 2).setValues(assignmentRows);
    }

    const createdTask = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdTask[header] = value instanceof Date ? value.toISOString() : value;
    });
    sendTaskAssignmentNotifications(createdTask, assigneeIds);

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

    const parsedDueDate = taskInput && taskInput.dueDate ? new Date(taskInput.dueDate) : '';
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const complexity = normalizeScaleValue(taskInput ? taskInput.complexity : '', 'Complexity');
    const priority = normalizeScaleValue(taskInput ? taskInput.priority : '', 'Priority');
    const description = taskInput && taskInput.description ? taskInput.description.toString().trim() : '';
    const assigneeIds = getValidAssigneeIds(taskInput ? taskInput.assigneeIds : []);
    const status = normalizeTaskStatus(taskInput ? taskInput.status : '');
    const projectId = ensureProjectExists(taskInput ? taskInput.projectId : '');

    if (headerIndex.Project_ID !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Project_ID + 1).setValue(projectId);
    if (headerIndex.Task_Title !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Task_Title + 1).setValue(taskTitle);
    if (headerIndex.Due_Date !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Due_Date + 1).setValue(hasValidDueDate ? parsedDueDate : '');
    if (headerIndex.Status !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Status + 1).setValue(status);
    if (headerIndex.Complexity !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Complexity + 1).setValue(complexity);
    if (headerIndex.Description !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Description + 1).setValue(description);
    if (headerIndex.Priority !== undefined) sheet.getRange(taskRowIndex + 1, headerIndex.Priority + 1).setValue(priority);

    const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
    if (!assignmentsSheet) {
      throw new Error('Assignments sheet was not found.');
    }

    const assignmentsData = assignmentsSheet.getDataRange().getValues();
    const previousAssigneeIds = assignmentsData
      .filter((row, index) => index > 0 && row[0] === normalizedTaskId)
      .map((row) => (row[1] || '').toString().trim())
      .filter(Boolean);
    for (let i = assignmentsData.length - 1; i >= 1; i--) {
      if (assignmentsData[i][0] === normalizedTaskId) {
        assignmentsSheet.deleteRow(i + 1);
      }
    }

    let updatedAssignments = [];
    if (assigneeIds.length > 0) {
      updatedAssignments = assigneeIds.map((assigneeId) => ({
        Assignment_ID: normalizedTaskId,
        Assignee_ID: assigneeId
      }));

      const assignmentRows = updatedAssignments.map((assignment) => [assignment.Assignment_ID, assignment.Assignee_ID]);
      assignmentsSheet.getRange(assignmentsSheet.getLastRow() + 1, 1, assignmentRows.length, 2).setValues(assignmentRows);
    }

    const refreshedData = sheet.getDataRange().getValues();
    const refreshedRow = refreshedData.find((row, index) => index > 0 && row[taskIdColumnIndex] === normalizedTaskId);

    const updatedTask = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      updatedTask[header] = value instanceof Date ? value.toISOString() : value;
    });
    const previousAssigneeSet = new Set(previousAssigneeIds);
    const newlyAssignedIds = assigneeIds.filter((assigneeId) => !previousAssigneeSet.has(assigneeId));
    sendTaskAssignmentNotifications(updatedTask, newlyAssignedIds);
    if (status === 'Completed' && previousStatus !== 'Completed') {
      const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
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
    const parsedDueDate = projectInput && projectInput.dueDate ? new Date(projectInput.dueDate) : '';
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';

    if (headerIndex.Project_Title !== undefined) {
      projectsSheet.getRange(projectRowIndex + 1, headerIndex.Project_Title + 1).setValue(projectTitle);
    }
    if (headerIndex.Description !== undefined) {
      projectsSheet.getRange(projectRowIndex + 1, headerIndex.Description + 1).setValue(description);
    }
    if (headerIndex.Status !== undefined) {
      projectsSheet.getRange(projectRowIndex + 1, headerIndex.Status + 1).setValue(status);
    }
    if (headerIndex.Due_Date !== undefined) {
      projectsSheet.getRange(projectRowIndex + 1, headerIndex.Due_Date + 1).setValue(hasValidDueDate ? parsedDueDate : '');
    }

    const refreshedData = projectsSheet.getDataRange().getValues();
    const refreshedRow = refreshedData.find((row, index) => index > 0 && row[projectIdColumnIndex] === normalizedProjectId);

    const updatedProject = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      updatedProject[header] = value instanceof Date ? value.toISOString() : value;
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
    for (let i = tasksData.length - 1; i >= 1; i--) {
      if (tasksData[i][taskProjectIdColumnIndex] === normalizedProjectId) {
        deletedTaskIds.push(tasksData[i][taskIdColumnIndex]);
        tasksSheet.deleteRow(i + 1);
      }
    }

    if (deletedTaskIds.length > 0) {
      const deletedTaskIdSet = new Set(deletedTaskIds);
      const assignmentsData = assignmentsSheet.getDataRange().getValues();
      for (let i = assignmentsData.length - 1; i >= 1; i--) {
        if (deletedTaskIdSet.has(assignmentsData[i][0])) {
          assignmentsSheet.deleteRow(i + 1);
        }
      }
    }

    projectsSheet.deleteRow(projectRowIndex + 1);

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
    const dueDateColumnIndex = headers.indexOf('Due_Date');

    if (taskIdColumnIndex === -1) throw new Error('Tasks sheet is missing Task_ID column.');
    if (statusColumnIndex === -1) throw new Error('Tasks sheet is missing Status column.');
    if (dueDateColumnIndex === -1) throw new Error('Tasks sheet is missing Due_Date column.');

    const taskRowIndex = taskData.findIndex((row, index) => index > 0 && row[taskIdColumnIndex] === normalizedTaskId);
    if (taskRowIndex < 0) throw new Error('Task not found.');

    const dueDateValue = taskData[taskRowIndex][dueDateColumnIndex];
    if (hasDueDatePassed(dueDateValue)) {
      deleteTaskAndAssignments(normalizedTaskId);
      return JSON.stringify({ success: true, deleted: true, taskId: normalizedTaskId });
    }

    tasksSheet.getRange(taskRowIndex + 1, statusColumnIndex + 1).setValue('Completed');
    const refreshedRow = tasksSheet.getRange(taskRowIndex + 1, 1, 1, headers.length).getValues()[0];
    const completedTask = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      completedTask[header] = value instanceof Date ? value.toISOString() : value;
    });
    const assignments = getTableData('Assignments');
    const assigneeIds = assignments
      .filter((assignment) => (assignment.Assignment_ID || '').toString().trim() === normalizedTaskId)
      .map((assignment) => (assignment.Assignee_ID || '').toString().trim())
      .filter(Boolean);
    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    sendManagerTaskCompletedNotifications(completedTask, assigneeIds, currentUserId);

    return JSON.stringify({ success: true, deleted: false, task: completedTask });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to complete task.'
    });
  }
}
