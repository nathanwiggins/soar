const USER_SETTINGS_PROPERTY_PREFIX = 'soar_user_settings:';
const DEFAULT_NOTIFICATION_SETTINGS = {
  taskAssignments: true,
  commentsAndMentions: true,
  dueDateReminders: true,
  weeklyDigest: false
};
function getDefaultNotificationSettings() {
  return Object.assign({}, DEFAULT_NOTIFICATION_SETTINGS);
}
function normalizeNotificationSettings(settingsInput) {
  const normalized = getDefaultNotificationSettings();
  const source = settingsInput && typeof settingsInput === 'object' ? settingsInput : {};
  Object.keys(normalized).forEach((key) => {
    if (typeof source[key] === 'boolean') {
      normalized[key] = source[key];
    }
  });
  return normalized;
}
function getUserSettingsPropertyKey(email) {
  const normalizedEmail = normalizeEmail(email);
  return normalizedEmail ? `${USER_SETTINGS_PROPERTY_PREFIX}${normalizedEmail}` : '';
}
function getUserSettingsByEmail(email, scriptPropertiesCache) {
  const propertyKey = getUserSettingsPropertyKey(email);
  if (!propertyKey) {
    return { notifications: getDefaultNotificationSettings() };
  }

  const rawSettings = scriptPropertiesCache
    ? scriptPropertiesCache[propertyKey]
    : PropertiesService.getScriptProperties().getProperty(propertyKey);
  if (!rawSettings) {
    return { notifications: getDefaultNotificationSettings() };
  }

  try {
    const parsedSettings = JSON.parse(rawSettings);
    return {
      notifications: normalizeNotificationSettings(parsedSettings && parsedSettings.notifications)
    };
  } catch (error) {
    Logger.log(`Failed to parse notification settings for ${normalizeEmail(email)}: ${error && error.message ? error.message : error}`);
    return { notifications: getDefaultNotificationSettings() };
  }
}
function isNotificationEnabledForEmail(email, notificationKey, scriptPropertiesCache) {
  const defaults = getDefaultNotificationSettings();
  if (!Object.prototype.hasOwnProperty.call(defaults, notificationKey)) return true;

  const settings = getUserSettingsByEmail(email, scriptPropertiesCache);
  return settings.notifications[notificationKey] !== false;
}
function isNotificationEnabledForUser(user, notificationKey, scriptPropertiesCache) {
  return isNotificationEnabledForEmail(user && user.Email, notificationKey, scriptPropertiesCache);
}
function persistCurrentUserSettings(settingsInput) {
  const currentUserEmail = normalizeEmail(getCurrentUser());
  if (!currentUserEmail) {
    return JSON.stringify({ success: false, error: 'Could not determine current user email.' });
  }

  const existingSettings = getUserSettingsByEmail(currentUserEmail);
  const requestedSettings = settingsInput && typeof settingsInput === 'object' ? settingsInput : {};
  const settingsToSave = {
    notifications: normalizeNotificationSettings(
      Object.prototype.hasOwnProperty.call(requestedSettings, 'notifications')
        ? requestedSettings.notifications
        : existingSettings.notifications
    )
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(getUserSettingsPropertyKey(currentUserEmail), JSON.stringify(settingsToSave));

  return JSON.stringify({ success: true, settings: settingsToSave });
}
