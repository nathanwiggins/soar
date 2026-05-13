function addSubtask(taskId, subtaskTitle) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
  if (!sheet) return JSON.stringify({ success: false, error: 'Subtasks sheet was not found.' });

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndex = getHeaderIndex(headers);
    const newRow = new Array(headers.length).fill('');
    
    if (headerIndex.Subtask_ID !== undefined) newRow[headerIndex.Subtask_ID] = generateNextId('Subtasks', 'S');
    if (headerIndex.Task_ID !== undefined) newRow[headerIndex.Task_ID] = taskId.toString().trim();
    if (headerIndex.Subtask_Title !== undefined) newRow[headerIndex.Subtask_Title] = subtaskTitle.toString().trim();
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = 'Incomplete';

    appendRows(sheet, [newRow]);
    invalidateTableCache('Subtasks');

    const createdSubtask = {};
    headers.forEach((header, index) => createdSubtask[header] = newRow[index]);

    return JSON.stringify({ success: true, subtask: createdSubtask });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to create subtask.' });
  }
}

function updateSubtaskStatus(subtaskId, isComplete) {
  const normalizedId = subtaskId ? subtaskId.toString().trim() : '';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
  
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('Subtask_ID');
    const statusCol = headers.indexOf('Status');
    
    const rowIndex = data.findIndex((row, i) => i > 0 && (row[idCol] || '').toString().trim() === normalizedId);
    if (rowIndex < 0) throw new Error('Subtask not found.');

    const updatedRow = data[rowIndex].slice();
    updatedRow[statusCol] = isComplete ? 'Complete' : 'Incomplete';
    
    updateRowValues(sheet, rowIndex + 1, updatedRow);
    invalidateTableCache('Subtasks');

    return JSON.stringify({ success: true });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message });
  }
}


function updateSubtaskTitle(subtaskId, subtaskTitle) {
  const normalizedId = subtaskId ? subtaskId.toString().trim() : '';
  const title = subtaskTitle ? subtaskTitle.toString().trim() : '';
  if (!normalizedId) return JSON.stringify({ success: false, error: 'Subtask ID is required.' });
  if (!title) return JSON.stringify({ success: false, error: 'Subtask title is required.' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
  if (!sheet) return JSON.stringify({ success: false, error: 'Subtasks sheet was not found.' });

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error('Subtasks sheet has no data rows.');

    const headers = data[0];
    const idCol = headers.indexOf('Subtask_ID');
    const titleCol = headers.indexOf('Subtask_Title');
    if (idCol === -1) throw new Error('Subtasks sheet is missing Subtask_ID column.');
    if (titleCol === -1) throw new Error('Subtasks sheet is missing Subtask_Title column.');

    const rowIndex = data.findIndex((row, i) => i > 0 && (row[idCol] || '').toString().trim() === normalizedId);
    if (rowIndex < 0) throw new Error('Subtask not found.');

    const updatedRow = data[rowIndex].slice();
    updatedRow[titleCol] = title;

    updateRowValues(sheet, rowIndex + 1, updatedRow);
    invalidateTableCache('Subtasks');

    const updatedSubtask = {};
    headers.forEach((header, index) => updatedSubtask[header] = updatedRow[index]);

    return JSON.stringify({ success: true, subtask: updatedSubtask });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to update subtask title.' });
  }
}

function reorderSubtasks(taskId, orderedSubtaskIds) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) return JSON.stringify({ success: false, error: 'Task ID is required.' });
  if (!Array.isArray(orderedSubtaskIds)) return JSON.stringify({ success: false, error: 'Subtask order is required.' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
  if (!sheet) return JSON.stringify({ success: false, error: 'Subtasks sheet was not found.' });

  try {
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) return JSON.stringify({ success: true, subtasks: [] });

    const headers = data[0];
    const idCol = headers.indexOf('Subtask_ID');
    const taskIdCol = headers.indexOf('Task_ID');
    if (idCol === -1) throw new Error('Subtasks sheet is missing Subtask_ID column.');
    if (taskIdCol === -1) throw new Error('Subtasks sheet is missing Task_ID column.');

    const normalizedOrderedIds = orderedSubtaskIds.map((id) => (id || '').toString().trim()).filter(Boolean);
    const uniqueOrderedIds = [...new Set(normalizedOrderedIds)];
    if (uniqueOrderedIds.length !== normalizedOrderedIds.length) {
      throw new Error('Subtask order contains duplicate IDs.');
    }

    const bodyRows = data.slice(1);
    const taskRows = bodyRows.filter((row) => (row[taskIdCol] || '').toString().trim() === normalizedTaskId);
    const existingIds = taskRows.map((row) => (row[idCol] || '').toString().trim()).filter(Boolean);
    const existingIdSet = new Set(existingIds);
    const orderedIdSet = new Set(uniqueOrderedIds);

    if (existingIds.length !== uniqueOrderedIds.length || existingIds.some((id) => !orderedIdSet.has(id)) || uniqueOrderedIds.some((id) => !existingIdSet.has(id))) {
      throw new Error('Subtask order must include every subtask for this task.');
    }

    const rowsById = taskRows.reduce((acc, row) => {
      acc[(row[idCol] || '').toString().trim()] = row;
      return acc;
    }, {});
    const reorderedTaskRows = uniqueOrderedIds.map((id) => rowsById[id]);
    let taskRowIndex = 0;
    const reorderedBodyRows = bodyRows.map((row) => {
      if ((row[taskIdCol] || '').toString().trim() !== normalizedTaskId) return row;
      return reorderedTaskRows[taskRowIndex++];
    });

    if (reorderedBodyRows.length > 0) {
      sheet.getRange(2, 1, reorderedBodyRows.length, headers.length).setValues(reorderedBodyRows);
    }
    invalidateTableCache('Subtasks');

    const subtasks = reorderedTaskRows.map((row) => {
      const subtask = {};
      headers.forEach((header, index) => subtask[header] = row[index]);
      return subtask;
    });

    return JSON.stringify({ success: true, subtasks });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to reorder subtasks.' });
  }
}

function deleteSubtask(subtaskId) {
  const normalizedId = subtaskId ? subtaskId.toString().trim() : '';
  if (!normalizedId) return JSON.stringify({ success: false, error: 'Subtask ID is required.' });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subtasks');
  if (!sheet) return JSON.stringify({ success: false, error: 'Subtasks sheet not found.' });

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) throw new Error('Subtasks sheet has no data rows.');

    const headers = data[0];
    const idCol = headers.indexOf('Subtask_ID');
    if (idCol === -1) throw new Error('Subtasks sheet is missing Subtask_ID column.');

    const rowIndex = data.findIndex((row, i) => i > 0 && (row[idCol] || '').toString().trim() === normalizedId);
    if (rowIndex < 0) throw new Error('Subtask not found.');

    deleteRowsBySheetIndexes(sheet, [rowIndex + 1]);
    invalidateTableCache('Subtasks');

    return JSON.stringify({ success: true, subtaskId: normalizedId });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to delete subtask.' });
  }
}