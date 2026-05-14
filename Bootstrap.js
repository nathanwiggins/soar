function getInitialPayload() {
  const currentUserEmail = normalizeEmail(getCurrentUser());
  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  let users = [];

  if (usersSheet) {
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0] || [];
    const headerIndex = getHeaderIndex(headers);
    const currentUserRow = currentUserEmail && headerIndex.Email !== undefined
      ? data.find((row, index) => index > 0 && normalizeEmail(row[headerIndex.Email]) === currentUserEmail)
      : null;
    const currentUserProfilePhotoUrl = currentUserRow && headerIndex.Profile_Pic_Url !== undefined
      ? (currentUserRow[headerIndex.Profile_Pic_Url] || '').toString().trim()
      : '';

    if (currentUserRow && headerIndex.Profile_Pic_Url !== undefined && !currentUserProfilePhotoUrl) {
      const fetchedProfilePhotoUrl = getCurrentUserProfilePhotoUrl();
      if (syncCurrentUserProfilePhoto(usersSheet, data, headerIndex, currentUserEmail, fetchedProfilePhotoUrl)) {
        invalidateTableCache('Users');
      }
    }
    users = getTableData('Users');
  }

  const currentUserExists = users.some((user) => normalizeEmail(user.Email) === currentUserEmail);
  const payload = {
    currentUserEmail: currentUserEmail,
    currentUserExists: currentUserExists,
    requiresAccountSetup: Boolean(currentUserEmail) && !currentUserExists,
    users: users,
    projects: getTableData('Projects'),
    tasks: getTableData('Tasks'),
    subtasks: getTableData('Subtasks'),
    assignments: getTableData('Assignments'),
    agendas: getTableData('Agendas'),
    agendaShares: getTableData('Sharing'),
    comments: getTableData('Comments').map((comment) => {
      const normalized = {};
      Object.keys(comment).forEach((key) => {
        normalized[key] = normalizeValueForClient(comment[key]);
      });
      return normalized;
    }),
    currentUserSettings: getUserSettingsByEmail(currentUserEmail),
    versionHash: buildGlobalVersionHash(),
    lastUpdated: getStoredDataVersion()
  };

  return JSON.stringify(payload);
}
