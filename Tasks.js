function getTaskById(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) return null;

  const tasks = getTableData('Tasks');
  return tasks.find((task) => (task.Task_ID || '').toString().trim() === normalizedTaskId) || null;
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

  return deletedTaskIds;
}
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

    const updatedTask = {};
    headers.forEach((header, index) => {
      const value = updatedRow[index];
      updatedTask[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });

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
    sendManagerTaskCompletedNotifications(completedTask, assigneeIds, currentUserId);

    return JSON.stringify({ success: true, task: completedTask });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to complete task.'
    });
  }
}
