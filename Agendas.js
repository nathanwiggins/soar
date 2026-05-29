function createAgenda(title, description) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Agendas');
  const creatorId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerIndex = getHeaderIndex(headers);
  const newRow = new Array(headers.length).fill('');

  if (headerIndex.Agenda_ID !== undefined) newRow[headerIndex.Agenda_ID] = generateNextId('Agendas', 'A');
  if (headerIndex.Title !== undefined) newRow[headerIndex.Title] = title.toString().trim();
  if (headerIndex.Description !== undefined) newRow[headerIndex.Description] = (description || '').toString().trim();
  if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;
  if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = new Date();

  appendRows(sheet, [newRow]);
  invalidateTableCache('Agendas');

  const createdAgenda = {};
  headers.forEach((header, index) => {
    const val = newRow[index];
    createdAgenda[header] = val instanceof Date ? val.toISOString() : val;
  });

  return JSON.stringify({ success: true, agenda: createdAgenda });
}

function createAgendaSession(agendaId, sessionDate, contentJson) {
  const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
  const agendasData = getTableData('Agendas');
  const agenda = agendasData.find(a => a.Agenda_ID === agendaId);
  if (!agenda) return JSON.stringify({ success: false, error: 'Agenda not found.' });

  const creatorId = (agenda.Creator_ID || '').toString().trim();
  if (creatorId !== (currentUserId || '').toString().trim()) {
    return JSON.stringify({ success: false, error: 'Only the agenda creator may create sessions.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  if (!sheet) return JSON.stringify({ success: false, error: 'Sessions sheet not found.' });

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerIndex = getHeaderIndex(headers);
  const newRow = new Array(headers.length).fill('');

  if (headerIndex.Session_ID !== undefined) newRow[headerIndex.Session_ID] = generateNextId('Sessions', 'AS');
  if (headerIndex.Agenda_ID !== undefined) newRow[headerIndex.Agenda_ID] = agendaId;
  if (headerIndex.Session_Date !== undefined) newRow[headerIndex.Session_Date] = sessionDate ? new Date(sessionDate + 'T00:00:00') : '';
  if (headerIndex.Content_JSON !== undefined) newRow[headerIndex.Content_JSON] = contentJson || '[]';
  if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = new Date();

  appendRows(sheet, [newRow]);
  invalidateTableCache('Sessions');

  const createdSession = {};
  headers.forEach((header, index) => {
    const val = newRow[index];
    createdSession[header] = val instanceof Date ? val.toISOString() : val;
  });

  return JSON.stringify({ success: true, session: createdSession });
}

function updateAgendaSession(sessionId, sessionDate, contentJson) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  if (!sheet) return JSON.stringify({ success: false, error: 'Sessions sheet not found.' });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const rowIndex = data.findIndex((row, i) => i > 0 && row[headerIndex.Session_ID] === sessionId);
  if (rowIndex < 0) return JSON.stringify({ success: false, error: 'Session not found.' });

  const updatedRow = data[rowIndex].slice();
  if (headerIndex.Session_Date !== undefined) updatedRow[headerIndex.Session_Date] = sessionDate ? new Date(sessionDate + 'T00:00:00') : '';
  if (headerIndex.Content_JSON !== undefined) updatedRow[headerIndex.Content_JSON] = contentJson || '[]';

  updateRowValues(sheet, rowIndex + 1, updatedRow);
  invalidateTableCache('Sessions');

  const updatedSession = {};
  headers.forEach((header, index) => {
    const val = updatedRow[index];
    updatedSession[header] = val instanceof Date ? val.toISOString() : val;
  });

  return JSON.stringify({ success: true, session: updatedSession });
}

function getAgendaSessions(agendaId) {
  const allSessions = getTableData('Sessions');
  const sessions = allSessions
    .filter(s => (s.Agenda_ID || '').toString().trim() === (agendaId || '').toString().trim())
    .map(s => {
      const normalized = {};
      Object.keys(s).forEach(key => { normalized[key] = normalizeValueForClient(s[key]); });
      return normalized;
    })
    .sort((a, b) => {
      const aDate = new Date(a.Session_Date || 0).getTime();
      const bDate = new Date(b.Session_Date || 0).getTime();
      return bDate - aDate;
    });
  return JSON.stringify({ success: true, sessions });
}

function deleteAgendaSession(sessionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  if (!sheet) return JSON.stringify({ success: false, error: 'Sessions sheet not found.' });

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const rowIndex = data.findIndex((row, i) => i > 0 && row[headerIndex.Session_ID] === sessionId);
  if (rowIndex < 0) return JSON.stringify({ success: false, error: 'Session not found.' });

  const agendaId = (data[rowIndex][headerIndex.Agenda_ID] || '').toString().trim();
  const agendasData = getTableData('Agendas');
  const agenda = agendasData.find(a => (a.Agenda_ID || '').toString().trim() === agendaId);
  if (agenda) {
    const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
    const creatorId = (agenda.Creator_ID || '').toString().trim();
    if (creatorId !== (currentUserId || '').toString().trim()) {
      return JSON.stringify({ success: false, error: 'Only the agenda creator can delete sessions.' });
    }
  }

  deleteRowsBySheetIndexes(sheet, [rowIndex + 1]);
  invalidateTableCache('Sessions');

  return JSON.stringify({ success: true, sessionId });
}

function migrateAgendaSessionsIfNeeded() {
  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  if (!sessionsSheet) return;

  const agendasData = getTableData('Agendas');
  const agendasWithContent = agendasData.filter(a => {
    const content = (a.Content_JSON || '').toString().trim();
    return content && content !== '[]' && content !== '';
  });

  if (agendasWithContent.length === 0) return;

  const existingSessions = getTableData('Sessions');
  const existingSessionAgendaIds = new Set(existingSessions.map(s => (s.Agenda_ID || '').toString().trim()));

  const headers = sessionsSheet.getRange(1, 1, 1, sessionsSheet.getLastColumn()).getValues()[0];
  const headerIndex = getHeaderIndex(headers);

  const rowsToAdd = [];
  agendasWithContent.forEach(agenda => {
    const agendaId = (agenda.Agenda_ID || '').toString().trim();
    if (!agendaId || existingSessionAgendaIds.has(agendaId)) return;

    const sessionDate = agenda.Agenda_Date || agenda.Created_Date || new Date();
    const dateVal = sessionDate instanceof Date ? sessionDate : new Date(sessionDate);
    const dateStr = `${dateVal.getFullYear()}-${String(dateVal.getMonth()+1).padStart(2,'0')}-${String(dateVal.getDate()).padStart(2,'0')}`;

    const newRow = new Array(headers.length).fill('');
    if (headerIndex.Session_ID !== undefined) newRow[headerIndex.Session_ID] = generateNextId('Sessions', 'AS');
    if (headerIndex.Agenda_ID !== undefined) newRow[headerIndex.Agenda_ID] = agendaId;
    if (headerIndex.Session_Date !== undefined) newRow[headerIndex.Session_Date] = new Date(dateStr + 'T00:00:00');
    if (headerIndex.Content_JSON !== undefined) newRow[headerIndex.Content_JSON] = agenda.Content_JSON || '[]';
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = new Date();
    rowsToAdd.push(newRow);
  });

  if (rowsToAdd.length > 0) {
    appendRows(sessionsSheet, rowsToAdd);
    invalidateTableCache('Sessions');
  }
}

function updateAgenda(agendaId, title, description, sharedUserIds) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Agendas');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const rowIndex = data.findIndex((row, i) => i > 0 && row[headerIndex.Agenda_ID] === agendaId);
  if (rowIndex < 0) return JSON.stringify({ success: false, error: 'Agenda not found.' });

  const updatedRow = data[rowIndex].slice();
  if (headerIndex.Title !== undefined) updatedRow[headerIndex.Title] = title;
  if (headerIndex.Description !== undefined) updatedRow[headerIndex.Description] = description || '';

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
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const rowIndex = data.findIndex((row, i) => i > 0 && row[0] === agendaId);
  if (rowIndex < 0) return JSON.stringify({ success: false, error: 'Agenda not found.' });

  const currentUserId = getCurrentUserIdByEmail(normalizeEmail(getCurrentUser()));
  const creatorId = (data[rowIndex][headerIndex.Creator_ID] || '').toString().trim();
  if (creatorId !== (currentUserId || '').toString().trim()) {
    return JSON.stringify({ success: false, error: 'Only the agenda creator can delete this agenda.' });
  }

  deleteRowsBySheetIndexes(sheet, [rowIndex + 1]);
  invalidateTableCache('Agendas');

  const sharesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sharing');
  const sharesData = sharesSheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = sharesData.length - 1; i >= 1; i--) {
    if (sharesData[i][0] === agendaId) rowsToDelete.push(i + 1);
  }
  deleteRowsBySheetIndexes(sharesSheet, rowsToDelete);
  invalidateTableCache('Sharing');

  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  if (sessionsSheet && sessionsSheet.getLastRow() > 1) {
    const sessionsData = sessionsSheet.getDataRange().getValues();
    const sessionsHeaders = sessionsData[0];
    const sessionsHeaderIndex = getHeaderIndex(sessionsHeaders);
    const sessionRowsToDelete = [];
    for (let i = sessionsData.length - 1; i >= 1; i--) {
      if (sessionsData[i][sessionsHeaderIndex.Agenda_ID] === agendaId) sessionRowsToDelete.push(i + 1);
    }
    deleteRowsBySheetIndexes(sessionsSheet, sessionRowsToDelete);
    invalidateTableCache('Sessions');
  }

  return JSON.stringify({ success: true, agendaId });
}