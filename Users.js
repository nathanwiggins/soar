function addUser(userInput) {
  const normalizedEmail = normalizeEmail(userInput && userInput.email);
  const name = userInput && userInput.name ? userInput.name.toString().trim() : '';
  const managerId = userInput && userInput.managerId ? userInput.managerId.toString().trim() : '';
  const currentUserEmail = normalizeEmail(getCurrentUser());
  const profilePicUrl = normalizedEmail === currentUserEmail ? getCurrentUserProfilePhotoUrl() : '';

  if (!normalizedEmail) {
    return JSON.stringify({ success: false, error: 'Email is required.' });
  }

  if (!isValidEmail(normalizedEmail)) {
    return JSON.stringify({ success: false, error: 'Email format is invalid.' });
  }

  if (!name) {
    return JSON.stringify({ success: false, error: 'Name is required.' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Users sheet was not found.' });
  }

  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const headerIndex = getHeaderIndex(headers);
    const emailColumnIndex = headerIndex.Email;
    const userIdColumnIndex = headerIndex.User_ID;

    if (emailColumnIndex === undefined || userIdColumnIndex === undefined) {
      throw new Error('Users sheet must contain User_ID and Email columns.');
    }

    const existingUser = data
      .slice(1)
      .find((row) => normalizeEmail(row[emailColumnIndex]) === normalizedEmail);

    if (existingUser) {
      const user = {};
      headers.forEach((header, index) => {
        user[header] = existingUser[index];
      });
      return JSON.stringify({ success: true, user: user, created: false });
    }

    if (managerId && headerIndex.Manager_ID !== undefined) {
      const users = getTableData('Users');
      const validManagerIds = new Set(users.map((user) => user.User_ID));
      if (!validManagerIds.has(managerId)) {
        throw new Error(`Manager_ID ${managerId} does not exist.`);
      }
    }

    const newRow = new Array(headers.length).fill('');
    if (headerIndex.User_ID !== undefined) newRow[headerIndex.User_ID] = generateNextId('Users', 'U');
    if (headerIndex.Email !== undefined) newRow[headerIndex.Email] = normalizedEmail;
    if (headerIndex.Name !== undefined) newRow[headerIndex.Name] = name;
    if (headerIndex.Manager_ID !== undefined) newRow[headerIndex.Manager_ID] = managerId;
    if (headerIndex.Profile_Pic_Url !== undefined) newRow[headerIndex.Profile_Pic_Url] = profilePicUrl;

    appendRows(sheet, [newRow]);
    invalidateTableCache('Users');

    const createdUser = {};
    headers.forEach((header, index) => {
      createdUser[header] = newRow[index];
    });
    sendManagerAccountCreatedNotification(createdUser);

    return JSON.stringify({ success: true, user: createdUser, created: true });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to add user.'
    });
  }
}
function getUserById(userId) {
  const normalizedUserId = userId ? userId.toString().trim() : '';
  if (!normalizedUserId) return null;

  const users = getTableData('Users');
  return users.find((user) => (user.User_ID || '').toString().trim() === normalizedUserId) || null;
}
function getUsersById() {
  return getTableData('Users').reduce((acc, user) => {
    const userId = (user.User_ID || '').toString().trim();
    if (userId) acc[userId] = user;
    return acc;
  }, {});
}
function updateCurrentUserProfile(profileInput) {
  const activeEmail = normalizeEmail(getCurrentUser());
  const normalizedName = profileInput && profileInput.name ? profileInput.name.toString().trim() : '';
  const normalizedEmail = normalizeEmail(profileInput ? profileInput.email : '');

  if (!normalizedName) {
    return JSON.stringify({ success: false, error: 'Name is required.' });
  }

  const usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!usersSheet) {
    return JSON.stringify({ success: false, error: 'Users sheet was not found.' });
  }

  try {
    const data = usersSheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Users sheet has no data rows.');
    }

    const headers = data[0];
    const headerIndex = getNormalizedHeaderIndex(headers);

    if (headerIndex.email === undefined) throw new Error('Users sheet is missing Email column.');
    if (headerIndex.name === undefined) throw new Error('Users sheet is missing Name column.');

    const currentUserRowIndex = data.findIndex(
      (row, index) => index > 0 && getUserEmailFromRow(row, headerIndex) === activeEmail
    );
    if (currentUserRowIndex < 0) {
      throw new Error('Current user record was not found.');
    }

    const updatedRow = data[currentUserRowIndex].slice();
    updatedRow[headerIndex.name] = normalizedName;

    if (headerIndex.profile_pic_url !== undefined) {
      const existingProfilePicUrl = updatedRow[headerIndex.profile_pic_url]
        ? updatedRow[headerIndex.profile_pic_url].toString().trim()
        : '';
      if (!existingProfilePicUrl) {
        const latestProfilePicUrl = getCurrentUserProfilePhotoUrl();
        if (latestProfilePicUrl) {
          updatedRow[headerIndex.profile_pic_url] = latestProfilePicUrl;
        }
      }
    }
    updateRowValues(usersSheet, currentUserRowIndex + 1, updatedRow);
    invalidateTableCache('Users');
    const updatedUser = {};
    headers.forEach((header, index) => {
      const value = updatedRow[index];
      updatedUser[header] = value instanceof Date ? value.toISOString() : value;
    });

    return JSON.stringify({ success: true, user: updatedUser });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to update user profile.'
    });
  }
}
function getCurrentUserIdByEmail(email) {
  if (!email) return '';

  const users = getTableData('Users');
  const normalizedEmail = email.toString().trim().toLowerCase();
  const user = users.find((u) => (u.Email || '').toString().trim().toLowerCase() === normalizedEmail);

  return user && user.User_ID ? user.User_ID : '';
}
function getAssignableUserIdsForUser(actorUserId) {
  const normalizedActorId = (actorUserId || '').toString().trim();
  if (!normalizedActorId) return new Set();

  const users = getTableData('Users');
  if (!Array.isArray(users) || users.length === 0) {
    return new Set([normalizedActorId]);
  }

  const directReportsByManager = users.reduce((acc, user) => {
    const managerId = (user.Manager_ID || '').toString().trim();
    const userId = (user.User_ID || '').toString().trim();
    if (!managerId || !userId) return acc;
    if (!acc[managerId]) acc[managerId] = [];
    acc[managerId].push(userId);
    return acc;
  }, {});

  const assignableIds = new Set([normalizedActorId]);
  const queue = [normalizedActorId];
  while (queue.length > 0) {
    const managerId = queue.shift();
    const reports = directReportsByManager[managerId] || [];
    reports.forEach((reportId) => {
      if (!assignableIds.has(reportId)) {
        assignableIds.add(reportId);
        queue.push(reportId);
      }
    });
  }

  return assignableIds;
}
