function normalizeProjectColorSchemeValue(value) {
  const validColorSchemes = ['red', 'blue', 'green', 'purple', 'amber', 'teal'];
  const normalizedValue = value ? value.toString().trim().toLowerCase() : 'red';
  return validColorSchemes.includes(normalizedValue) ? normalizedValue : 'red';
}
function getProjectHeadersWithColorScheme(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (headers.indexOf('Color_Scheme') !== -1) return headers;

  const nextColumn = headers.length + 1;
  sheet.getRange(1, nextColumn).setValue('Color_Scheme');
  return headers.concat('Color_Scheme');
}
function getProjectById(projectId) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) return null;

  const projects = getTableData('Projects');
  return projects.find((project) => (project.Project_ID || '').toString().trim() === normalizedProjectId) || null;
}
function createProject(projectInput) {
  const projectTitle = projectInput && projectInput.projectTitle ? projectInput.projectTitle.toString().trim() : '';
  if (!projectTitle) {
    return JSON.stringify({ success: false, error: 'Project title is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Projects sheet was not found.' });
  }

  try {
    const headers = getProjectHeadersWithColorScheme(sheet);
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    const newRow = new Array(headers.length).fill('');
    const now = new Date();
    const status = normalizeStatusValue(projectInput ? projectInput.status : '');
    const description = projectInput && projectInput.description ? projectInput.description.toString().trim() : '';
    const parsedDueDate = parseDateInput(projectInput ? projectInput.dueDate : '');
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const colorScheme = normalizeProjectColorSchemeValue(projectInput ? projectInput.colorScheme : '');
    const creatorId = getCurrentUserIdByEmail(getCurrentUser());

    if (!creatorId) {
      throw new Error('Could not determine Creator_ID from current user email.');
    }

    if (headerIndex.Project_ID !== undefined) newRow[headerIndex.Project_ID] = generateNextId('Projects', 'P');
    if (headerIndex.Project_Title !== undefined) newRow[headerIndex.Project_Title] = projectTitle;
    if (headerIndex.Description !== undefined) newRow[headerIndex.Description] = description;
    if (headerIndex.Status !== undefined) newRow[headerIndex.Status] = status;
    if (headerIndex.Created_Date !== undefined) newRow[headerIndex.Created_Date] = now;
    if (headerIndex.Due_Date !== undefined) newRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Creator_ID !== undefined) newRow[headerIndex.Creator_ID] = creatorId;
    if (headerIndex.Color_Scheme !== undefined) newRow[headerIndex.Color_Scheme] = colorScheme;

    appendRows(sheet, [newRow]);
    invalidateTableCache('Projects');

    const projectId = headerIndex.Project_ID !== undefined ? newRow[headerIndex.Project_ID] : '';
    if (projectId) {
      const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Assignments');
      if (!assignmentsSheet) {
        throw new Error('Assignments sheet was not found.');
      }
      appendRows(assignmentsSheet, [[projectId, creatorId]]);
      invalidateTableCache('Assignments');
    }

    const createdProject = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdProject[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });

    const createdAssignment = projectId
      ? { Assignment_ID: projectId, Assignee_ID: creatorId }
      : null;

    return JSON.stringify({ success: true, project: createdProject, assignment: createdAssignment });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to create project.'
    });
  }
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
function updateProject(projectId, projectInput) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) {
    return JSON.stringify({ success: false, error: 'Project ID is required.' });
  }

  const projectTitle = projectInput && projectInput.projectTitle ? projectInput.projectTitle.toString().trim() : '';
  if (!projectTitle) {
    return JSON.stringify({ success: false, error: 'Project title is required.' });
  }

  const projectsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projects');
  if (!projectsSheet) {
    return JSON.stringify({ success: false, error: 'Projects sheet was not found.' });
  }

  try {
    getProjectHeadersWithColorScheme(projectsSheet);
    const dataRange = projectsSheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      throw new Error('Projects sheet has no data rows.');
    }

    const headers = data[0];
    const headerIndex = headers.reduce((acc, header, index) => {
      acc[header] = index;
      return acc;
    }, {});

    const projectIdColumnIndex = headerIndex.Project_ID;
    if (projectIdColumnIndex === undefined) {
      throw new Error('Projects sheet is missing Project_ID column.');
    }

    const projectRowIndex = data.findIndex((row, index) => index > 0 && row[projectIdColumnIndex] === normalizedProjectId);
    if (projectRowIndex < 0) {
      throw new Error('Project not found.');
    }

    const status = normalizeStatusValue(projectInput ? projectInput.status : '');
    const description = projectInput && projectInput.description ? projectInput.description.toString().trim() : '';
    const parsedDueDate = parseDateInput(projectInput ? projectInput.dueDate : '');
    const hasValidDueDate = parsedDueDate && parsedDueDate.toString() !== 'Invalid Date';
    const colorScheme = normalizeProjectColorSchemeValue(projectInput ? projectInput.colorScheme : '');

    const refreshedRow = data[projectRowIndex].slice();
    if (headerIndex.Project_Title !== undefined) refreshedRow[headerIndex.Project_Title] = projectTitle;
    if (headerIndex.Description !== undefined) refreshedRow[headerIndex.Description] = description;
    if (headerIndex.Status !== undefined) refreshedRow[headerIndex.Status] = status;
    if (headerIndex.Due_Date !== undefined) refreshedRow[headerIndex.Due_Date] = hasValidDueDate ? parsedDueDate : '';
    if (headerIndex.Color_Scheme !== undefined) refreshedRow[headerIndex.Color_Scheme] = colorScheme;
    updateRowValues(projectsSheet, projectRowIndex + 1, refreshedRow);
    invalidateTableCache('Projects');

    const updatedProject = {};
    headers.forEach((header, index) => {
      const value = refreshedRow[index];
      updatedProject[header] = header === 'Due_Date'
        ? serializeDateOnlyForClient(value)
        : (value instanceof Date ? value.toISOString() : value);
    });

    return JSON.stringify({ success: true, project: updatedProject });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update project.'
    });
  }
}
function deleteProject(projectId) {
  const normalizedProjectId = projectId ? projectId.toString().trim() : '';
  if (!normalizedProjectId) {
    return JSON.stringify({ success: false, error: 'Project ID is required.' });
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const projectsSheet = spreadsheet.getSheetByName('Projects');
  const tasksSheet = spreadsheet.getSheetByName('Tasks');
  const assignmentsSheet = spreadsheet.getSheetByName('Assignments');

  if (!projectsSheet) {
    return JSON.stringify({ success: false, error: 'Projects sheet was not found.' });
  }
  if (!tasksSheet) {
    return JSON.stringify({ success: false, error: 'Tasks sheet was not found.' });
  }
  if (!assignmentsSheet) {
    return JSON.stringify({ success: false, error: 'Assignments sheet was not found.' });
  }

  try {
    const projectData = projectsSheet.getDataRange().getValues();
    if (projectData.length <= 1) {
      throw new Error('Projects sheet has no data rows.');
    }

    const projectHeaders = projectData[0];
    const projectIdColumnIndex = projectHeaders.indexOf('Project_ID');
    if (projectIdColumnIndex === -1) {
      throw new Error('Projects sheet is missing Project_ID column.');
    }

    const projectRowIndex = projectData.findIndex(
      (row, index) => index > 0 && row[projectIdColumnIndex] === normalizedProjectId
    );
    if (projectRowIndex < 0) {
      throw new Error('Project not found.');
    }

    const tasksData = tasksSheet.getDataRange().getValues();
    const taskHeaders = tasksData[0] || [];
    const taskIdColumnIndex = taskHeaders.indexOf('Task_ID');
    const taskProjectIdColumnIndex = taskHeaders.indexOf('Project_ID');
    if (taskIdColumnIndex === -1 || taskProjectIdColumnIndex === -1) {
      throw new Error('Tasks sheet must contain Task_ID and Project_ID columns.');
    }

    const deletedTaskIds = [];
    const taskRowsToDelete = [];
    for (let i = tasksData.length - 1; i >= 1; i--) {
      if (tasksData[i][taskProjectIdColumnIndex] === normalizedProjectId) {
        deletedTaskIds.push(tasksData[i][taskIdColumnIndex]);
        taskRowsToDelete.push(i + 1);
      }
    }
    deleteRowsBySheetIndexes(tasksSheet, taskRowsToDelete);
    invalidateTableCache('Tasks');

    if (deletedTaskIds.length > 0) {
      const deletedTaskIdSet = new Set(deletedTaskIds);
      const assignmentsData = assignmentsSheet.getDataRange().getValues();
      const assignmentRowsToDelete = [];
      for (let i = assignmentsData.length - 1; i >= 1; i--) {
        if (deletedTaskIdSet.has(assignmentsData[i][0])) {
          assignmentRowsToDelete.push(i + 1);
        }
      }
      deleteRowsBySheetIndexes(assignmentsSheet, assignmentRowsToDelete);
      invalidateTableCache('Assignments');
    }

    deleteRowsBySheetIndexes(projectsSheet, [projectRowIndex + 1]);
    invalidateTableCache('Projects');

    return JSON.stringify({
      success: true,
      projectId: normalizedProjectId,
      deletedTaskIds: deletedTaskIds
    });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to delete project.'
    });
  }
}
