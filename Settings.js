const USER_SETTINGS_PROPERTY_PREFIX = 'soar_user_settings:';
const USER_SORT_ORDER_PREFIX = 'soar_sort:';
const DEFAULT_NOTIFICATION_SETTINGS = {
  taskAssignments: true,
  commentsAndMentions: true,
  dueDateReminders: true,
  taskCompletion: true,
  weeklyDigest: false,
  agendaShares: true,
  ticketFollowUp: true
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
function getUserSortOrderPropertyKey(email, entityType) {
  const normalizedEmail = normalizeEmail(email);
  return normalizedEmail ? `${USER_SORT_ORDER_PREFIX}${normalizedEmail}:${entityType}` : '';
}

function getUserSortOrders(email, scriptPropertiesCache) {
  const projectsKey = getUserSortOrderPropertyKey(email, 'projects');
  const tasksKey = getUserSortOrderPropertyKey(email, 'tasks');
  if (!projectsKey) return { projects: [], tasks: [] };

  const cache = scriptPropertiesCache || PropertiesService.getScriptProperties().getProperties();
  let projects = [];
  let tasks = [];
  try { if (cache[projectsKey]) projects = JSON.parse(cache[projectsKey]); } catch (e) {}
  try { if (cache[tasksKey]) tasks = JSON.parse(cache[tasksKey]); } catch (e) {}
  return { projects, tasks };
}

function saveUserSortOrder(entityType, orderedIds) {
  const currentUserEmail = normalizeEmail(getCurrentUser());
  if (!currentUserEmail) return JSON.stringify({ success: false, error: 'Could not determine current user.' });
  if (entityType !== 'projects' && entityType !== 'tasks') return JSON.stringify({ success: false, error: 'Invalid entity type.' });
  if (!Array.isArray(orderedIds)) return JSON.stringify({ success: false, error: 'orderedIds must be an array.' });
  const key = getUserSortOrderPropertyKey(currentUserEmail, entityType);
  if (!key) return JSON.stringify({ success: false, error: 'Invalid user email.' });
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(orderedIds));
  return JSON.stringify({ success: true });
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
