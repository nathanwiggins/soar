function createAgenda(title) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Agendas');
  const creatorId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerIndex = getHeaderIndex(headers);
  const newRow = new Array(headers.length).fill('');
  
  if (headerIndex.Agenda_ID !== undefined) newRow[headerIndex.Agenda_ID] = generateNextId('Agendas', 'A');
  if (headerIndex.Title !== undefined) newRow[headerIndex.Title] = title.toString().trim();
  if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;
  if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = new Date();
  if (headerIndex.Content_JSON !== undefined) newRow[headerIndex.Content_JSON] = '[]';

  appendRows(sheet, [newRow]);
  invalidateTableCache('Agendas');

  const createdAgenda = {};
  headers.forEach((header, index) => {
    const val = newRow[index];
    createdAgenda[header] = val instanceof Date ? val.toISOString() : val;
  });

  return JSON.stringify({ success: true, agenda: createdAgenda });
}

function updateAgenda(agendaId, title, description, agendaDate, contentJson, sharedUserIds) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Agendas');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const rowIndex = data.findIndex((row, i) => i > 0 && row[headerIndex.Agenda_ID] === agendaId);
  if (rowIndex < 0) return JSON.stringify({ success: false, error: 'Agenda not found.' });

  const updatedRow = data[rowIndex].slice();
  if (headerIndex.Title !== undefined) updatedRow[headerIndex.Title] = title;
  if (headerIndex.Description !== undefined) updatedRow[headerIndex.Description] = description || '';
  if (headerIndex.Agenda_Date !== undefined) updatedRow[headerIndex.Agenda_Date] = agendaDate ? new Date(agendaDate + 'T00:00:00') : '';
  if (headerIndex.Content_JSON !== undefined) updatedRow[headerIndex.Content_JSON] = contentJson;
  
  updateRowValues(sheet, rowIndex + 1, updatedRow);
  invalidateTableCache('Agendas');

  const sharesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sharing');
  const sharesData = sharesSheet.getDataRange().getValues();
  const previousSharedUserIds = [];
  const rowsToDelete = [];
  for (let i = sharesData.length - 1; i >= 1; i--) {
    if (sharesData[i][0] === agendaId) {
      previousSharedUserIds.push((sharesData[i][1] || '').toString().trim());
      rowsToDelete.push(i + 1);
    }
  }
  deleteRowsBySheetIndexes(sharesSheet, rowsToDelete);
  
  let newShares = [];
  if (sharedUserIds && sharedUserIds.length > 0) {
    newShares = sharedUserIds.map(uid => [agendaId, uid]);
    appendRows(sharesSheet, newShares);
  }
  invalidateTableCache('Sharing');

  const updatedAgenda = {};
  headers.forEach((header, index) => {
    const val = updatedRow[index];
    updatedAgenda[header] = val instanceof Date ? val.toISOString() : val;
  });

  const previousSharesSet = new Set(previousSharedUserIds);
  const newlySharedIds = (sharedUserIds || []).filter(uid => !previousSharesSet.has(uid.toString().trim()));
  const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
  
  if (typeof sendAgendaShareNotifications === 'function') {
    sendAgendaShareNotifications(updatedAgenda, newlySharedIds, currentUserId);
  }

  return JSON.stringify({ 
    success: true, 
    agenda: updatedAgenda, 
    shares: newShares.map(s => ({ Agenda_ID: s[0], User_ID: s[1] })) 
  });
}

function deleteAgenda(agendaId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Agendas');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, i) => i > 0 && row[0] === agendaId);
  if (rowIndex > -1) {
    deleteRowsBySheetIndexes(sheet, [rowIndex + 1]);
    invalidateTableCache('Agendas');
  }

  const sharesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sharing');
  const sharesData = sharesSheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = sharesData.length - 1; i >= 1; i--) {
    if (sharesData[i][0] === agendaId) rowsToDelete.push(i + 1);
  }
  deleteRowsBySheetIndexes(sharesSheet, rowsToDelete);
  invalidateTableCache('Sharing');

  return JSON.stringify({ success: true, agendaId });
}