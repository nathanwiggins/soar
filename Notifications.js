function getTaskNotificationDetails(task, projectsById) {
  const taskTitle = (task && task.Task_Title ? task.Task_Title : '').toString().trim() || 'Untitled task';
  const projectId = (task && task.Project_ID ? task.Project_ID : '').toString().trim();
  const project = projectsById
    ? projectsById[projectId]
    : getProjectById(projectId);
  const projectTitle = (project && project.Project_Title ? project.Project_Title : '').toString().trim() || 'Unassigned project';

  return {
    taskTitle,
    projectTitle
  };
}
function formatTaskNotificationBody(details) {
  const taskTitle = details && details.taskTitle ? details.taskTitle : 'Untitled task';
  const projectTitle = details && details.projectTitle ? details.projectTitle : 'Unassigned project';
  return `Task: ${taskTitle}\nProject: ${projectTitle}`;
}
function ensureTaskExists(taskId) {
  const normalizedTaskId = taskId ? taskId.toString().trim() : '';
  if (!normalizedTaskId) {
    throw new Error('Topic_ID (task ID) is required.');
  }

  const tasks = getTableData('Tasks');
  const taskExists = tasks.some((task) => (task.Task_ID || '').toString().trim() === normalizedTaskId);
  if (!taskExists) {
    throw new Error(`Task ${normalizedTaskId} does not exist.`);
  }

  return normalizedTaskId;
}
function sendMentionNotifications(comment, topicId, commenter, mentionedUsers) {
  if (!mentionedUsers || mentionedUsers.length === 0) return;

  const task = getTaskById(topicId);
  const details = getTaskNotificationDetails(task);
  const commenterName = commenter && commenter.Name ? commenter.Name : 'A teammate';
  const subject = `You were mentioned on ${details.taskTitle}`;
  const body = `${commenterName} mentioned you in a comment.

${formatTaskNotificationBody(details)}

Comment:
${comment}
`;

  const sentEmails = new Set();
  mentionedUsers.forEach((user) => {
    const email = normalizeEmail(user.Email);
    if (!email || sentEmails.has(email)) return;
    if (!isNotificationEnabledForUser(user, 'commentsAndMentions')) return;
    sentEmails.add(email);
    safeSendEmail(email, subject, body);
  });
}
function sendTaskAssignmentNotifications(task, assigneeIds, assignedByUserId) {
  if (!task || !Array.isArray(assigneeIds) || assigneeIds.length === 0) return;

  const usersById = getUsersById();
  const assignedBy = usersById[(assignedByUserId || '').toString().trim()];
  const assignedByName = assignedBy && assignedBy.Name ? assignedBy.Name : 'A teammate';
  const details = getTaskNotificationDetails(task);
  const subject = `New task assignment: ${details.taskTitle}`;

  const sentEmails = new Set();
  assigneeIds.forEach((assigneeId) => {
    const user = usersById[(assigneeId || '').toString().trim()];
    const email = normalizeEmail(user && user.Email);
    if (!email || sentEmails.has(email)) return;
    if (!isNotificationEnabledForUser(user, 'taskAssignments')) return;

    const recipientName = user && user.Name ? user.Name : 'there';
    const body = `Hi ${recipientName},

${assignedByName} assigned you to a task.

${formatTaskNotificationBody(details)}
`;

    sentEmails.add(email);
    safeSendEmail(email, subject, body);
  });
}
function sendManagerTaskCompletedNotifications(task, assigneeIds, completedByUserId) {
  if (!task || !Array.isArray(assigneeIds) || assigneeIds.length === 0) return;

  const users = getTableData('Users');
  const usersById = users.reduce((acc, user) => {
    const userId = (user.User_ID || '').toString().trim();
    if (userId) acc[userId] = user;
    return acc;
  }, {});

  const managerIds = new Set();
  assigneeIds.forEach((assigneeId) => {
    const worker = usersById[(assigneeId || '').toString().trim()];
    const managerId = (worker && worker.Manager_ID ? worker.Manager_ID : '').toString().trim();
    if (managerId) managerIds.add(managerId);
  });
  if (managerIds.size === 0) return;

  const details = getTaskNotificationDetails(task);
  const completedBy = usersById[(completedByUserId || '').toString().trim()];
  const completedByName = completedBy && completedBy.Name ? completedBy.Name : 'A user';
  const subject = `Task completed: ${details.taskTitle}`;
  const body = `${completedByName} marked a task as completed.

${formatTaskNotificationBody(details)}
`;

  const sentEmails = new Set();
  managerIds.forEach((managerId) => {
    const manager = usersById[managerId];
    const email = normalizeEmail(manager && manager.Email);
    if (!email || sentEmails.has(email)) return;
    sentEmails.add(email);
    safeSendEmail(email, subject, body);
  });
}
function sendManagerAccountCreatedNotification(createdUser) {
  if (!createdUser) return;
  const managerId = (createdUser.Manager_ID || '').toString().trim();
  if (!managerId) return;

  const manager = getUserById(managerId);
  const managerEmail = normalizeEmail(manager && manager.Email);
  if (!managerEmail) return;

  const userId = (createdUser.User_ID || '').toString().trim();
  const userName = (createdUser.Name || '').toString().trim() || userId || 'A user';
  const userEmail = normalizeEmail(createdUser.Email);
  const subject = `New direct report created: ${userName}`;
  const body = `${userName}${userEmail ? ` (${userEmail})` : ''} created an account and listed you as their manager.\n\nUser ID: ${userId || 'N/A'}\n`;
  safeSendEmail(managerEmail, subject, body);
}
function hasDueDatePassed(dueDateValue) {
  if (!dueDateValue) return false;
  const dueDate = new Date(dueDateValue);
  if (dueDate.toString() === 'Invalid Date') return false;

  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const normalizedDueDate = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());

  return normalizedDueDate.getTime() < startOfToday.getTime();
}
function getDaysUntilDate(dateValue) {
  if (!dateValue) return null;
  const dueDate = parseDateInput(dateValue);
  if (!dueDate || dueDate.toString() === 'Invalid Date') return null;

  const today = new Date();
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const normalizedDueDate = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate());
  return Math.round((normalizedDueDate.getTime() - startOfToday.getTime()) / (24 * 60 * 60 * 1000));
}
function getAssignmentsByAssignmentId() {
  return getTableData('Assignments').reduce((acc, assignment) => {
    const assignmentId = (assignment.Assignment_ID || '').toString().trim();
    const assigneeId = (assignment.Assignee_ID || '').toString().trim();
    if (!assignmentId || !assigneeId) return acc;
    if (!acc[assignmentId]) acc[assignmentId] = [];
    acc[assignmentId].push(assigneeId);
    return acc;
  }, {});
}
function sendDueDateReminderNotifications() {
  const tasks = getTableData('Tasks');
  const assignmentsByTaskId = getAssignmentsByAssignmentId();
  const usersById = getUsersById();
  const scriptProperties = PropertiesService.getScriptProperties();
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let sentCount = 0;

  tasks.forEach((task) => {
    const status = (task.Status || '').toString().trim();
    if (status === 'Completed') return;

    const daysUntilDue = getDaysUntilDate(task.Due_Date);
    if (daysUntilDue === null || daysUntilDue < 0 || daysUntilDue > 1) return;

    const taskId = (task.Task_ID || '').toString().trim();
    const details = getTaskNotificationDetails(task);
    const dueLabel = daysUntilDue === 0 ? 'today' : 'tomorrow';
    const subject = `Task due ${dueLabel}: ${details.taskTitle}`;
    const body = `A task is due ${dueLabel}.

${formatTaskNotificationBody(details)}
`;
    const sentEmails = new Set();

    (assignmentsByTaskId[taskId] || []).forEach((assigneeId) => {
      const user = usersById[assigneeId];
      const email = normalizeEmail(user && user.Email);
      if (!email || sentEmails.has(email)) return;
      if (!isNotificationEnabledForUser(user, 'dueDateReminders')) return;

      const reminderKey = `soar_due_date_reminder:${todayKey}:${taskId}:${email}`;
      if (scriptProperties.getProperty(reminderKey)) return;

      sentEmails.add(email);
      if (safeSendEmail(email, subject, body)) {
        scriptProperties.setProperty(reminderKey, new Date().toISOString());
        sentCount += 1;
      }
    });
  });

  return sentCount;
}
function sendWeeklyDigestNotifications() {
  const tasks = getTableData('Tasks');
  const assignmentsByTaskId = getAssignmentsByAssignmentId();
  const usersById = getUsersById();
  const tasksByUserId = {};

  tasks.forEach((task) => {
    const status = (task.Status || '').toString().trim();
    if (status === 'Completed') return;

    const taskId = (task.Task_ID || '').toString().trim();
    (assignmentsByTaskId[taskId] || []).forEach((assigneeId) => {
      if (!tasksByUserId[assigneeId]) tasksByUserId[assigneeId] = [];
      tasksByUserId[assigneeId].push(task);
    });
  });

  const digestDateKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const scriptProperties = PropertiesService.getScriptProperties();
  let sentCount = 0;

  Object.keys(tasksByUserId).forEach((userId) => {
    const user = usersById[userId];
    const email = normalizeEmail(user && user.Email);
    if (!email || !isNotificationEnabledForUser(user, 'weeklyDigest')) return;

    const digestKey = `soar_weekly_digest:${digestDateKey}:${email}`;
    if (scriptProperties.getProperty(digestKey)) return;

    const taskLines = tasksByUserId[userId]
      .sort((left, right) => {
        const leftDays = getDaysUntilDate(left.Due_Date);
        const rightDays = getDaysUntilDate(right.Due_Date);
        return (leftDays === null ? 9999 : leftDays) - (rightDays === null ? 9999 : rightDays);
      })
      .map((task) => {
        const details = getTaskNotificationDetails(task);
        const dueDate = task.Due_Date ? serializeDateOnlyForClient(parseDateInput(task.Due_Date)) : 'No due date';
        return `- ${details.taskTitle} (${details.projectTitle}), due: ${dueDate}`;
      });

    const subject = 'Your weekly Soar task digest';
    const body = 'Here are your open Soar tasks for this week:\n\n' + taskLines.join('\n') + '\n';
    if (safeSendEmail(email, subject, body)) {
      scriptProperties.setProperty(digestKey, new Date().toISOString());
      sentCount += 1;
    }
  });

  return sentCount;
}
