function getInitialPayload() {
  const currentUserEmail = normalizeEmail(getCurrentUser());
  const currentUserProfilePhotoUrl = getCurrentUserProfilePhotoUrl();
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  let users = [];

  if (usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const headerIndex = getHeaderIndex(headers);

    if (syncCurrentUserProfilePhoto(usersSheet, data, headerIndex, currentUserEmail, currentUserProfilePhotoUrl)) {
      invalidateTableCache('Users');
    }
    users = getTableData('Users');
  }

  const currentUserExists = users.some((user) => normalizeEmail(user.Email) === currentUserEmail);
  const completionMetadataByTaskId = getTaskCompletionMetadataMap();
  const payload = {
    currentUserEmail: currentUserEmail,
    currentUserExists: currentUserExists,
    requiresAccountSetup: Boolean(currentUserEmail) && !currentUserExists,
    users: users,
    projects: getTableData('Projects'),
    tasks: getTableData('Tasks').map((task) => appendTaskCompletionMetadataToTask(task, completionMetadataByTaskId)),
    assignments: getTableData('Assignments'),
    currentUserSettings: getUserSettingsByEmail(currentUserEmail),
    versionHash: buildGlobalVersionHash(),
    lastUpdated: getStoredDataVersion()
  };

  return JSON.stringify(payload);
}
