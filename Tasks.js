const TASK_DEFAULT_STATUS = 'Not Started';
const TASK_COMPLETE_STATUS = 'Complete';
const TASK_STATUS_OPTIONS = [
  TASK_DEFAULT_STATUS,
  'Upcoming',
  'Review',
  'In Progress',
  'Ongoing',
  'On Hold',
  'Cancelled',
  'Closeout',
  TASK_COMPLETE_STATUS
];
const RECURRENCE_RULE_OPTIONS = ['', 'Daily', 'Weekly', 'Monthly', 'Yearly'];
function isTaskCompleteStatus(status) {
  return ['Complete', 'Completed'].includes((status || '').toString().trim());
}
function normalizeLegacyTaskStatus(status) {
  const normalizedStatus = status ? status.toString().trim() : '';
  if (normalizedStatus === 'Not Started') return TASK_DEFAULT_STATUS;
  return isTaskCompleteStatus(normalizedStatus) ? TASK_COMPLETE_STATUS : normalizedStatus;
}
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
    if (isTaskCompleteStatus(status) && hasDueDatePassed(dueDateValue)) {
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
    const isMarkingCompleted = isTaskCompleteStatus(status) && !isTaskCompleteStatus(previousStatus);
    const isReopeningCompleted = !isTaskCompleteStatus(status) && isTaskCompleteStatus(previousStatus);
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
      sendTaskCompletedNotifications(updatedTask, assigneeIds, currentUserId);
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
function normalizeRecurrenceRule(value) {
  const rule = value ? value.toString().trim() : '';
  if (rule === 'None') return '';
  if (RECURRENCE_RULE_OPTIONS.indexOf(rule) === -1) {
    throw new Error(`Recurrence must be one of: ${RECURRENCE_RULE_OPTIONS.filter(Boolean).join(', ')}.`);
  }
  return rule;
}
function normalizeRecurrenceInterval(value, rule) {
  if (!rule) return '';

  const interval = value === null || value === undefined || value === '' ? 1 : parseInt(value, 10);
  if (!Number.isInteger(interval) || interval < 1) {
    throw new Error('Recurrence interval must be a whole number of 1 or more.');
  }
  return interval;
}
function normalizeRecurrenceEndDate(value, rule) {
  if (!rule) return '';

  const parsedEndDate = parseDateInput(value);
  return parsedEndDate && parsedEndDate.toString() !== 'Invalid Date' ? parsedEndDate : '';
}
function daysInMonth(year, month) {
  return new Date(year, month + 1, 0).getDate();
}
function computeNextRecurrenceDueDate(baseDate, rule, interval) {
  const year = baseDate.getFullYear();
  const month = baseDate.getMonth();
  const day = baseDate.getDate();

  if (rule === 'Daily') return new Date(year, month, day + interval);
  if (rule === 'Weekly') return new Date(year, month, day + interval * 7);
  if (rule === 'Monthly' || rule === 'Yearly') {
    const normalized = rule === 'Yearly' ? new Date(year + interval, month, 1) : new Date(year, month + interval, 1);
    const targetYear = normalized.getFullYear();
    const targetMonth = normalized.getMonth();
    const isLastDayOfBaseMonth = day >= daysInMonth(year, month);
    const targetDay = isLastDayOfBaseMonth ? daysInMonth(targetYear, targetMonth) : Math.min(day, daysInMonth(targetYear, targetMonth));
    return new Date(targetYear, targetMonth, targetDay);
  }
  return new Date(year, month, day);
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
    const recurrenceRule = normalizeRecurrenceRule(taskInput ? taskInput.recurrenceRule : '');
    const recurrenceInterval = normalizeRecurrenceInterval(taskInput ? taskInput.recurrenceInterval : '', recurrenceRule);
    const recurrenceEndDate = normalizeRecurrenceEndDate(taskInput ? taskInput.recurrenceEndDate : '', recurrenceRule);
    if (recurrenceRule && !hasValidDueDate) {
      throw new Error('A due date is required to set up a recurring task.');
    }
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
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = TASK_DEFAULT_STATUS;
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = now;
    if (headerIndex.Description !== undefined) newRow[headerIndex.Description] = description;
    if (headerIndex.Priority !== undefined) newRow[headerIndex.Priority] = priority;
    if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;
    if (headerIndex.Recurrence_Rule !== undefined) newRow[headerIndex.Recurrence_Rule] = recurrenceRule;
    if (headerIndex.Recurrence_Interval !== undefined) newRow[headerIndex.Recurrence_Interval] = recurrenceInterval;
    if (headerIndex.Recurrence_End_Date !== undefined) newRow[headerIndex.Recurrence_End_Date] = recurrenceEndDate;

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
      createdTask[header] = header === 'Due_Date' || header === 'Recurrence_End_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });
    sendTaskAssignmentNotifications(createdTask, assigneeIds, creatorId);

    const subtaskTitles = Array.isArray(taskInput?.subtasks) ? taskInput.subtasks : [];
    let createdSubtasks = [];
    
    if (subtaskTitles.length > 0 && headerIndex.Task_ID !== undefined) {
      const subtasksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
      if (subtasksSheet) {
        const subtaskHeaders = subtasksSheet.getRange(1, 1, 1, subtasksSheet.getLastColumn()).getValues()[0];
        const subHeaderIdx = getHeaderIndex(subtaskHeaders);
        
        const subtaskRows = subtaskTitles.map(title => {
          const row = new Array(subtaskHeaders.length).fill('');
          if (subHeaderIdx.Subtask_ID !== undefined) row[subHeaderIdx.Subtask_ID] = generateNextId('Subtasks', 'S');
          if (subHeaderIdx.Task_ID !== undefined) row[subHeaderIdx.Task_ID] = newRow[headerIndex.Task_ID];
          if (subHeaderIdx.Subtask_Title !== undefined) row[subHeaderIdx.Subtask_Title] = title.toString().trim();
          if (subHeaderIdx.Status !== undefined) row[subHeaderIdx.Status] = 'Incomplete';
          return row;
        });
        
        appendRows(subtasksSheet, subtaskRows);
        invalidateTableCache('Subtasks');
        
        createdSubtasks = subtaskRows.map(row => {
            const obj = {};
            subtaskHeaders.forEach((h, i) => obj[h] = row[i]);
            return obj;
        });
      }
    }

    return JSON.stringify({ success: true, task: createdTask, assignments: createdAssignments, subtasks: createdSubtasks });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to create task.'
    });
  }
}
function normalizeTaskStatus(status) {
  const allowedStatuses = TASK_STATUS_OPTIONS;
  const normalizedStatus = normalizeLegacyTaskStatus(status);

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
    const recurrenceRule = normalizeRecurrenceRule(taskInput ? taskInput.recurrenceRule : '');
    const recurrenceInterval = normalizeRecurrenceInterval(taskInput ? taskInput.recurrenceInterval : '', recurrenceRule);
    const recurrenceEndDate = normalizeRecurrenceEndDate(taskInput ? taskInput.recurrenceEndDate : '', recurrenceRule);
    if (recurrenceRule && !hasValidDueDate) {
      throw new Error('A due date is required to set up a recurring task.');
    }
    const assigneeIds = getValidAssigneeIds(taskInput ? taskInput.assigneeIds : []);
    ensureTaskHasAssignees(assigneeIds);
    const status = normalizeTaskStatus(taskInput ? taskInput.status : '');
    const projectId = ensureProjectExists(taskInput ? taskInput.projectId : '');
    const editorId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));

    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    const isMarkingCompleted = isTaskCompleteStatus(status) && !isTaskCompleteStatus(previousStatus);
    const isReopeningCompleted = !isTaskCompleteStatus(status) && isTaskCompleteStatus(previousStatus);

    const updatedTaskRow = data[taskRowIndex].slice();
    if (headerIndex.Project_ID !== undefined) updatedTaskRow[headerIndex.Project_ID] = projectId;
    if (headerIndex.Task_Title !== undefined) updatedTaskRow[headerIndex.Task_Title] = taskTitle;
    if (headerIndex.Due_Date !== undefined) updatedTaskRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Status !== undefined) updatedTaskRow[headerIndex.Status] = status;
    if (headerIndex.Description !== undefined) updatedTaskRow[headerIndex.Description] = description;
    if (headerIndex.Priority !== undefined) updatedTaskRow[headerIndex.Priority] = priority;
    if (headerIndex.Recurrence_Rule !== undefined) updatedTaskRow[headerIndex.Recurrence_Rule] = recurrenceRule;
    if (headerIndex.Recurrence_Interval !== undefined) updatedTaskRow[headerIndex.Recurrence_Interval] = recurrenceInterval;
    if (headerIndex.Recurrence_End_Date !== undefined) updatedTaskRow[headerIndex.Recurrence_End_Date] = recurrenceEndDate;
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
      updatedTask[header] = header === 'Due_Date' || header === 'Recurrence_End_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });
    const previousAssigneeSet = new Set(previousAssigneeIds);
    const newlyAssignedIds = assigneeIds.filter((assigneeId) => !previousAssigneeSet.has(assigneeId));
    sendTaskAssignmentNotifications(updatedTask, newlyAssignedIds, currentUserId);
    if (isMarkingCompleted) {
      sendTaskCompletedNotifications(updatedTask, assigneeIds, currentUserId);
    }

    const subtaskTitles = Array.isArray(taskInput?.newSubtasks) ? taskInput.newSubtasks : [];
    let createdSubtasks = [];

    if (subtaskTitles.length > 0 && headerIndex.Task_ID !== undefined) {
      const subtasksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
      if (subtasksSheet) {
        const subtaskHeaders = subtasksSheet.getRange(1, 1, 1, subtasksSheet.getLastColumn()).getValues()[0];
        const subHeaderIdx = getHeaderIndex(subtaskHeaders);

        const subtaskRows = subtaskTitles.map(title => {
          const row = new Array(subtaskHeaders.length).fill('');
          if (subHeaderIdx.Subtask_ID !== undefined) row[subHeaderIdx.Subtask_ID] = generateNextId('Subtasks', 'S');
          if (subHeaderIdx.Task_ID !== undefined) row[subHeaderIdx.Task_ID] = normalizedTaskId;
          if (subHeaderIdx.Subtask_Title !== undefined) row[subHeaderIdx.Subtask_Title] = title.toString().trim();
          if (subHeaderIdx.Status !== undefined) row[subHeaderIdx.Status] = 'Incomplete';
          return row;
        });

        appendRows(subtasksSheet, subtaskRows);
        invalidateTableCache('Subtasks');

        createdSubtasks = subtaskRows.map(row => {
            const obj = {};
            subtaskHeaders.forEach((h, i) => obj[h] = row[i]);
            return obj;
        });
      }
    }

    return JSON.stringify({ success: true, task: updatedTask, assignments: updatedAssignments, newSubtasks: createdSubtasks });
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
    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    const taskData = getTableData('Tasks');
    const task = taskData.find(t => (t.Task_ID || '').toString().trim() === normalizedTaskId);
    if (!task) return JSON.stringify({ success: false, error: 'Task not found.' });
    const taskCreatorId = (task.Creator_ID || '').toString().trim();
    if (taskCreatorId && taskCreatorId !== (currentUserId || '').toString().trim()) {
      return JSON.stringify({ success: false, error: 'Only the task creator can delete this task.' });
    }

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
    refreshedRow[statusColumnIndex] = TASK_COMPLETE_STATUS;
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
    sendTaskCompletedNotifications(completedTask, assigneeIds, currentUserId);

    return JSON.stringify({ success: true, task: completedTask });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to complete task.'
    });
  }
}

function moveTaskToProject(taskId, newProjectId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  const normalizedProjectId = newProjectId ? newProjectId.toString().trim() : '';
  if (!normalizedTaskId || !normalizedProjectId) return JSON.stringify({ success: false, error: 'IDs missing.' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sheet) return JSON.stringify({ success: false, error: 'Tasks sheet not found.' });

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify({ success: true });

    const headers = data[0];
    const idCol = headers.indexOf('Task_ID');
    const projCol = headers.indexOf('Project_ID');
    const bodyRows = data.slice(1);

    const taskRow = bodyRows.find(row => (row[idCol] || '').toString().trim() === normalizedTaskId);
    if (!taskRow) return JSON.stringify({ success: false, error: 'Task not found.' });

    const currentProjectId = (taskRow[projCol] || '').toString().trim();
    if (currentProjectId !== normalizedProjectId) {
      const currentUserEmail = normalizeEmail(getCurrentUser());
      const users = getTableData('Users');
      const currentUserRecord = users.find(u => normalizeEmail(u.Email) === currentUserEmail);
      const currentUserId = currentUserRecord ? (currentUserRecord.User_ID || '').toString().trim() : '';
      const projects = getTableData('Projects');
      const sourceProject = projects.find(p => (p.Project_ID || '').toString().trim() === currentProjectId);
      const targetProject = projects.find(p => (p.Project_ID || '').toString().trim() === normalizedProjectId);
      const sourceCreatorId = (sourceProject?.Creator_ID || '').toString().trim();
      const targetCreatorId = (targetProject?.Creator_ID || '').toString().trim();
      if (!currentUserId || currentUserId !== sourceCreatorId || currentUserId !== targetCreatorId) {
        return JSON.stringify({ success: false, error: 'Only project creators can move tasks between projects.' });
      }
    }

    const taskRowIndex = bodyRows.findIndex(row => (row[idCol] || '').toString().trim() === normalizedTaskId);
    if (taskRowIndex >= 0) {
      bodyRows[taskRowIndex][projCol] = normalizedProjectId;
      sheet.getRange(2, 1, bodyRows.length, headers.length).setValues(bodyRows);
      invalidateTableCache('Tasks');
    }

    return JSON.stringify({ success: true });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message });
  }
}

function updateTaskDueDate(taskId, dueDate) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    return JSON.stringify({ success: false, error: 'Task ID is required.' });
  }

  const parsedDueDate = parseDateInput(dueDate);
  if (!parsedDueDate || parsedDueDate.toString() === 'Invalid Date') {
    return JSON.stringify({ success: false, error: 'A valid due date is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error('Tasks sheet has no data rows.');

    const headers = data[0];
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    if (headerIndex.Task_ID === undefined || headerIndex.Due_Date === undefined) {
      throw new Error('Tasks sheet is missing required columns.');
    }

    const rowIndex = data.findIndex((row, index) => index > 0 && row[headerIndex.Task_ID] === normalizedTaskId);
    if (rowIndex < 0) throw new Error('Task not found.');

    const updatedRow = data[rowIndex].slice();
    updatedRow[headerIndex.Due_Date] = parsedDueDate;
    updateRowValues(sheet, rowIndex + 1, updatedRow);
    invalidateTableCache('Tasks');

    const updatedTask = {};
    headers.forEach((header, index) => {
      const value = updatedRow[index];
      updatedTask[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });

    return JSON.stringify({ success: true, task: updatedTask });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update task due date.'
    });
  }
}

function processRecurringTasks() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Tasks');
  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;

  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);
  if (headerIndex.Recurrence_Rule === undefined || headerIndex.Due_Date === undefined || headerIndex.Task_ID === undefined) {
    return 0;
  }

  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const assignmentsByTaskId = getAssignmentsByAssignmentId();
  const subtasksByTaskId = getTableData('Subtasks').reduce((acc, subtask) => {
    const taskId = (subtask.Task_ID || '').toString().trim();
    if (!taskId) return acc;
    if (!acc[taskId]) acc[taskId] = [];
    acc[taskId].push(subtask);
    return acc;
  }, {});

  const rowIndexesToClear = [];
  const spawnedTasks = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rule = (row[headerIndex.Recurrence_Rule] || '').toString().trim();
    if (!rule || RECURRENCE_RULE_OPTIONS.indexOf(rule) === -1) continue;

    const dueDate = parseDateInput(row[headerIndex.Due_Date]);
    if (!dueDate || dueDate.toString() === 'Invalid Date' || dueDate.getTime() > startOfToday.getTime()) continue;

    const rawInterval = parseInt(row[headerIndex.Recurrence_Interval], 10);
    const interval = Number.isInteger(rawInterval) && rawInterval >= 1 ? rawInterval : 1;
    const endDate = headerIndex.Recurrence_End_Date !== undefined
      ? parseDateInput(row[headerIndex.Recurrence_End_Date])
      : '';

    let nextDue = computeNextRecurrenceDueDate(dueDate, rule, interval);
    while (nextDue.getTime() <= startOfToday.getTime()) {
      nextDue = computeNextRecurrenceDueDate(nextDue, rule, interval);
    }

    rowIndexesToClear.push(i);

    if (endDate && endDate.toString() !== 'Invalid Date' && nextDue.getTime() > endDate.getTime()) {
      continue;
    }

    const newRow = row.slice();
    const newTaskId = generateNextId('Tasks', 'T');
    newRow[headerIndex.Task_ID] = newTaskId;
    newRow[headerIndex.Due_Date] = nextDue;
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = TASK_DEFAULT_STATUS;
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = new Date();
    if (headerIndex.Completed_By !== undefined) newRow[headerIndex.Completed_By] = '';
    if (headerIndex.Completed_At !== undefined) newRow[headerIndex.Completed_At] = '';

    const sourceTaskId = (row[headerIndex.Task_ID] || '').toString().trim();
    spawnedTasks.push({
      newRow,
      newTaskId,
      creatorId: headerIndex.Creator_ID !== undefined ? (row[headerIndex.Creator_ID] || '').toString().trim() : '',
      assigneeIds: assignmentsByTaskId[sourceTaskId] || [],
      subtasks: subtasksByTaskId[sourceTaskId] || []
    });
  }

  if (rowIndexesToClear.length === 0) return 0;

  rowIndexesToClear.forEach((rowIndex) => {
    if (headerIndex.Recurrence_Rule !== undefined) sheet.getRange(rowIndex + 1, headerIndex.Recurrence_Rule + 1).setValue('');
    if (headerIndex.Recurrence_Interval !== undefined) sheet.getRange(rowIndex + 1, headerIndex.Recurrence_Interval + 1).setValue('');
    if (headerIndex.Recurrence_End_Date !== undefined) sheet.getRange(rowIndex + 1, headerIndex.Recurrence_End_Date + 1).setValue('');
  });

  if (spawnedTasks.length === 0) {
    invalidateTableCache('Tasks');
    return 0;
  }

  appendRows(sheet, spawnedTasks.map((spawned) => spawned.newRow));
  invalidateTableCache('Tasks');

  const assignmentsSheet = spreadsheet.getSheetByName('Assignments');
  const subtasksSheet = spreadsheet.getSheetByName('Subtasks');
  const subtaskHeaders = subtasksSheet ? subtasksSheet.getRange(1, 1, 1, subtasksSheet.getLastColumn()).getValues()[0] : [];
  const subtaskHeaderIndex = getHeaderIndex(subtaskHeaders);

  const assignmentRows = [];
  const subtaskRows = [];

  spawnedTasks.forEach((spawned) => {
    spawned.assigneeIds.forEach((assigneeId) => {
      assignmentRows.push([spawned.newTaskId, assigneeId]);
    });

    if (subtasksSheet) {
      spawned.subtasks.forEach((subtask) => {
        const subtaskRow = new Array(subtaskHeaders.length).fill('');
        if (subtaskHeaderIndex.Subtask_ID !== undefined) subtaskRow[subtaskHeaderIndex.Subtask_ID] = generateNextId('Subtasks', 'S');
        if (subtaskHeaderIndex.Task_ID !== undefined) subtaskRow[subtaskHeaderIndex.Task_ID] = spawned.newTaskId;
        if (subtaskHeaderIndex.Subtask_Title !== undefined) subtaskRow[subtaskHeaderIndex.Subtask_Title] = subtask.Subtask_Title || '';
        if (subtaskHeaderIndex.Status !== undefined) subtaskRow[subtaskHeaderIndex.Status] = 'Incomplete';
        subtaskRows.push(subtaskRow);
      });
    }
  });

  if (assignmentsSheet && assignmentRows.length > 0) {
    appendRows(assignmentsSheet, assignmentRows);
    invalidateTableCache('Assignments');
  }
  if (subtasksSheet && subtaskRows.length > 0) {
    appendRows(subtasksSheet, subtaskRows);
    invalidateTableCache('Subtasks');
  }

  spawnedTasks.forEach((spawned) => {
    const newTaskObject = {};
    headers.forEach((header, index) => {
      const value = spawned.newRow[index];
      newTaskObject[header] = header === 'Due_Date' || header === 'Recurrence_End_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });
    sendTaskAssignmentNotifications(newTaskObject, spawned.assigneeIds, spawned.creatorId);
  });

  return spawnedTasks.length;
}
