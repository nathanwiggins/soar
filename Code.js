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

function normalizeScaleValue(value, fieldName) {
  if (value === null || value === undefined || value === '') return '';

  const parsedValue = Number(value);
  if (!Number.isInteger(parsedValue) || parsedValue < 1 || parsedValue > 5) {
    throw new Error(`${fieldName} must be an integer between 1 and 5.`);
  }
  return parsedValue;
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

    return JSON.stringify({ success: true, task: updatedTask, assignments: updatedAssignments });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update task.'
    });
  }
}
