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