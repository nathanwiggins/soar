/**
 * Main entry point for the web app.
 * Serves the Index.html file.
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
  
  // Assuming ID is always the first column based on Data Dictionary
  const lastIdStr = data[data.length - 1][0]; 
  const numericPart = parseInt(lastIdStr.split('-')[1], 10);
  const nextNumber = numericPart + 1;
  
  return `${prefix}-${nextNumber.toString().padStart(8, '0')}`;
}

/**
 * API Endpoint: Fetches the full data payload for the frontend to initialize.
 */
function getInitialPayload() {
  const payload = {
    currentUserEmail: getCurrentUser(),
    users: getTableData('Users'),
    projects: getTableData('Projects'),
    tasks: getTableData('Tasks'),
    assignments: getTableData('Assignments')
  };
  
  // Stringifying prevents Apps Script's silent serialization failures
  return JSON.stringify(payload); 
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

/**
 * API Endpoint: Creates a task for a project.
 */
function createTask(projectId, taskTitle) {
  if (!projectId) {
    return { success: false, error: 'Project ID is required.' };
  }

  if (!taskTitle || !taskTitle.toString().trim()) {
    return { success: false, error: 'Task title is required.' };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sheet) {
    return { success: false, error: 'Tasks sheet was not found.' };
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerIndex = headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});

  const newRow = new Array(headers.length).fill('');
  const now = new Date();

  if (headerIndex.Task_ID !== undefined) newRow[headerIndex.Task_ID] = generateNextId('Tasks', 'T');
  if (headerIndex.Project_ID !== undefined) newRow[headerIndex.Project_ID] = projectId;
  if (headerIndex.Task_Title !== undefined) newRow[headerIndex.Task_Title] = taskTitle.toString().trim();
  if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = 'Not Started';
  if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = now;

  sheet.appendRow(newRow);

  const createdTask = {};
  headers.forEach((header, index) => {
    createdTask[header] = newRow[index];
  });

  return { success: true, task: createdTask };
}
