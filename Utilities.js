function getCurrentUser() {
  return Session.getActiveUser().getEmail();
}
function getCurrentUserProfilePhotoUrl() {
  const activeEmail = normalizeEmail(getCurrentUser());
  if (!activeEmail) return '';

  try {
    const response = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      return '';
    }

    const data = JSON.parse(response.getContentText() || '{}');
    const profileEmail = normalizeEmail(data.email);
    if (profileEmail && profileEmail !== activeEmail) {
      return '';
    }

    return data.picture ? data.picture.toString().trim() : '';
  } catch (error) {
    return '';
  }
}
function syncCurrentUserProfilePhoto(sheet, data, headerIndex, currentUserEmail, profilePicUrl) {
  if (!sheet || !data || !headerIndex) return false;
  if (!currentUserEmail || !profilePicUrl) return false;

  const emailColumnIndex = headerIndex.Email;
  const profilePicColumnIndex = headerIndex.Profile_Pic_Url;

  if (emailColumnIndex === undefined || profilePicColumnIndex === undefined) {
    return false;
  }

  const matchingRowIndex = data.findIndex(
    (row, index) => index > 0 && normalizeEmail(row[emailColumnIndex]) === currentUserEmail
  );

  if (matchingRowIndex < 0) {
    return false;
  }

  const existingProfilePicUrl = data[matchingRowIndex][profilePicColumnIndex]
    ? data[matchingRowIndex][profilePicColumnIndex].toString().trim()
    : '';

  if (existingProfilePicUrl === profilePicUrl) {
    return false;
  }

  try {
    sheet.getRange(matchingRowIndex + 1, profilePicColumnIndex + 1).setValue(profilePicUrl);
    return true;
  } catch (error) {

    return false;
  }
}
function normalizeEmail(email) {
  return email ? email.toString().trim().toLowerCase() : '';
}
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}
function getHeaderIndex(headers) {
  return headers.reduce((acc, header, index) => {
    acc[header] = index;
    return acc;
  }, {});
}
function normalizeHeaderName(header) {
  return header ? header.toString().trim().toLowerCase() : '';
}
function getNormalizedHeaderIndex(headers) {
  return headers.reduce((acc, header, index) => {
    acc[normalizeHeaderName(header)] = index;
    return acc;
  }, {});
}
function getUserEmailFromRow(row, headerIndex) {
  if (!row || !headerIndex) return '';
  const emailColumnIndex = headerIndex.email;
  if (emailColumnIndex === undefined) return '';
  return normalizeEmail(row[emailColumnIndex]);
}
function safeSendEmail(recipient, subject, body) {
  const normalizedRecipient = normalizeEmail(recipient);
  if (!normalizedRecipient) return false;

  try {
    MailApp.sendEmail(normalizedRecipient, subject || '', body || '');
    return true;
  } catch (error) {

    Logger.log(`Failed to send notification email to ${normalizedRecipient}: ${error && error.message ? error.message : error}`);
    return false;
  }
}
function parseDateInput(dateInput) {
  if (!dateInput) return '';

  if (Object.prototype.toString.call(dateInput) === '[object Date]') {
    if (Number.isNaN(dateInput.getTime())) return '';
    return new Date(dateInput.getFullYear(), dateInput.getMonth(), dateInput.getDate());
  }

  const trimmedValue = dateInput.toString().trim();
  if (!trimmedValue) return '';

  const dateOnlyMatch = trimmedValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateOnlyMatch) {
    const year = Number(dateOnlyMatch[1]);
    const month = Number(dateOnlyMatch[2]) - 1;
    const day = Number(dateOnlyMatch[3]);
    return new Date(year, month, day);
  }

  const parsedDate = new Date(trimmedValue);
  if (parsedDate.toString() === 'Invalid Date') return '';
  return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
}
function serializeDateOnlyForClient(dateValue) {
  if (!(Object.prototype.toString.call(dateValue) === '[object Date]') || Number.isNaN(dateValue.getTime())) {
    return dateValue;
  }
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
