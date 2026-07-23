function runDailyTriggers() {
  const jobs = [
    processRecurringTasks,
    syncDailyGitHubStatus,
    sendDueDateReminderNotifications,
    purgeCompletedTasksPastDue,
  ];

  jobs.forEach((job) => {
    try {
      job();
    } catch (error) {
      Logger.log(`runDailyTriggers: ${job.name} failed: ${error}`);
    }
  });
}

function runWeeklyTriggers() {
  const jobs = [sendWeeklyDigestNotifications];

  jobs.forEach((job) => {
    try {
      job();
    } catch (error) {
      Logger.log(`runWeeklyTriggers: ${job.name} failed: ${error}`);
    }
  });
}

function runMonthlyTriggers() {
  const jobs = [];

  jobs.forEach((job) => {
    try {
      job();
    } catch (error) {
      Logger.log(`runMonthlyTriggers: ${job.name} failed: ${error}`);
    }
  });
}
