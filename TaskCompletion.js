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
