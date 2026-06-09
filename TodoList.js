const TODO_PROPERTY_PREFIX = 'soar_todo:';

function getTodoPropertyKey(email) {
  const normalizedEmail = normalizeEmail(email);
  return normalizedEmail ? `${TODO_PROPERTY_PREFIX}${normalizedEmail}` : '';
}

function readTodoItems(email) {
  const key = getTodoPropertyKey(email);
  if (!key) return [];
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function writeTodoItems(email, items) {
  const key = getTodoPropertyKey(email);
  if (!key) return;
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(items));
}

function generateTodoId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const key = 'soar_todo_counter';
    const props = PropertiesService.getScriptProperties();
    const counter = parseInt(props.getProperty(key) || '0', 10) + 1;
    props.setProperty(key, counter.toString());
    return 'todo-' + counter;
  } finally {
    lock.releaseLock();
  }
}

function getTodoList() {
  const email = normalizeEmail(getCurrentUser());
  if (!email) return JSON.stringify({ success: false, error: 'Not authenticated.' });
  return JSON.stringify({ success: true, items: readTodoItems(email) });
}

function saveTodoList(itemsJson) {
  const email = normalizeEmail(getCurrentUser());
  if (!email) return JSON.stringify({ success: false, error: 'Not authenticated.' });
  let items;
  try {
    items = typeof itemsJson === 'string' ? JSON.parse(itemsJson) : itemsJson;
    if (!Array.isArray(items)) throw new Error('Not an array.');
  } catch (e) {
    return JSON.stringify({ success: false, error: 'Invalid items data.' });
  }
  writeTodoItems(email, items);
  return JSON.stringify({ success: true });
}

function addLinkedTodoItem(type, linkedId, label) {
  const email = normalizeEmail(getCurrentUser());
  if (!email) return JSON.stringify({ success: false, error: 'Not authenticated.' });
  const validTypes = ['project', 'task', 'subtask'];
  if (!validTypes.includes(type)) return JSON.stringify({ success: false, error: 'Invalid type.' });
  const normalizedLinkedId = linkedId ? linkedId.toString().trim() : '';
  const normalizedLabel = label ? label.toString().trim() : '';
  if (!normalizedLinkedId || !normalizedLabel) return JSON.stringify({ success: false, error: 'Missing linkedId or label.' });

  const items = readTodoItems(email);
  const alreadyLinked = items.some((item) => item.type === type && item.linkedId === normalizedLinkedId);
  if (alreadyLinked) return JSON.stringify({ success: true, items, alreadyLinked: true });

  const newItem = {
    id: generateTodoId(),
    type,
    label: normalizedLabel,
    linkedId: normalizedLinkedId,
    completed: false
  };
  items.push(newItem);
  writeTodoItems(email, items);
  return JSON.stringify({ success: true, items });
}

function toggleTodoItemComplete(itemId, completed) {
  const email = normalizeEmail(getCurrentUser());
  if (!email) return JSON.stringify({ success: false, error: 'Not authenticated.' });

  const items = readTodoItems(email);
  const item = items.find((i) => i.id === itemId);
  if (!item) return JSON.stringify({ success: false, error: 'Item not found.' });

  item.completed = Boolean(completed);

  if (item.linkedId) {
    if (item.type === 'task') {
      updateTaskStatus(item.linkedId, completed ? 'Complete' : 'Not Started');
    } else if (item.type === 'project') {
      updateProjectStatusForTodo(item.linkedId, completed ? 'Completed' : 'Not Started');
    } else if (item.type === 'subtask') {
      updateSubtaskStatus(item.linkedId, Boolean(completed));
    }
  }

  writeTodoItems(email, items);
  return JSON.stringify({ success: true, items });
}

function updateProjectStatusForTodo(projectId, status) {
  const normalizedId = projectId ? projectId.toString().trim() : '';
  if (!normalizedId) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if (!sheet) return;
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    const headers = data[0];
    const idCol = headers.indexOf('Project_ID');
    const statusCol = headers.indexOf('Status');
    if (idCol === -1 || statusCol === -1) return;
    const rowIndex = data.findIndex((row, i) => i > 0 && (row[idCol] || '').toString().trim() === normalizedId);
    if (rowIndex < 0) return;
    const updatedRow = data[rowIndex].slice();
    updatedRow[statusCol] = status;
    updateRowValues(sheet, rowIndex + 1, updatedRow);
    invalidateTableCache('Projects');
  } catch (e) {
    Logger.log('updateProjectStatusForTodo error: ' + (e.message || e));
  }
}

function removeTodoItem(itemId) {
  const email = normalizeEmail(getCurrentUser());
  if (!email) return JSON.stringify({ success: false, error: 'Not authenticated.' });
  const items = readTodoItems(email).filter((i) => i.id !== itemId);
  writeTodoItems(email, items);
  return JSON.stringify({ success: true, items });
}

function resetTodoList() {
  const email = normalizeEmail(getCurrentUser());
  if (!email) return JSON.stringify({ success: false, error: 'Not authenticated.' });
  writeTodoItems(email, []);
  return JSON.stringify({ success: true, items: [] });
}
