function logSupportTicket(issueSummary) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Issues sheet not found.' });
  }

  const issueId = generateNextId('Issues', 'ISSUE');
  const timestamp = new Date().toISOString();
  const userEmail = normalizeEmail(getCurrentUser());

  const githubPat = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
  let status = 'New';
  let githubIssueNumber = '';

  if (githubPat) {
    const githubResponse = UrlFetchApp.fetch('https://api.github.com/repos/nathanwiggins/soar/issues', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${githubPat}`, 'User-Agent': 'SOAR-App' },
      payload: JSON.stringify({
        title: `[User Report] ${issueSummary.substring(0, 100)}`,
        body: `**Reported by:** ${userEmail}\n**Timestamp:** ${timestamp}\n\n${issueSummary}`
      }),
      muteHttpExceptions: true
    });

    if (githubResponse.getResponseCode() === 201) {
      const githubIssue = JSON.parse(githubResponse.getContentText());
      status = 'Pending';
      githubIssueNumber = githubIssue.number;
    } else {
      Logger.log(`logSupportTicket: GitHub API error: ${githubResponse.getContentText()}`);
    }
  }

  appendRows(sheet, [[issueId, timestamp, userEmail, issueSummary, status, githubIssueNumber]]);
  invalidateTableCache('Issues');

  return JSON.stringify({ success: true, issueId: issueId });
}

function syncDailyGitHubStatus() {
  const githubPat = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');
  if (!githubPat) {
    Logger.log('syncDailyGitHubStatus: Missing GITHUB_PAT');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const pendingRows = data.slice(1)
    .map((row, i) => ({ row, sheetRowIndex: i + 2 }))
    .filter(({ row }) => row[headerIndex.Status] === 'Pending' && row[headerIndex.GitHub_Issue_Number]);

  pendingRows.forEach(({ row, sheetRowIndex }) => {
    const issueNumber = row[headerIndex.GitHub_Issue_Number];
    const githubResponse = UrlFetchApp.fetch(`https://api.github.com/repos/nathanwiggins/soar/issues/${issueNumber}`, {
      headers: { Authorization: `Bearer ${githubPat}`, 'User-Agent': 'SOAR-App' },
      muteHttpExceptions: true
    });

    if (githubResponse.getResponseCode() !== 200) {
      Logger.log(`syncDailyGitHubStatus: GitHub API error for issue #${issueNumber}: ${githubResponse.getContentText()}`);
      return;
    }

    const githubIssue = JSON.parse(githubResponse.getContentText());
    if (githubIssue.state === 'closed') {
      const newRow = [...row];
      newRow[headerIndex.Status] = 'Complete';
      updateRowValues(sheet, sheetRowIndex, newRow);

      safeSendEmail(
        row[headerIndex.User_Email],
        'Your SOAR issue has been resolved',
        `Great news! The issue you reported regarding '${row[headerIndex.Issue_Description]}' has been resolved by our development team. Thank you for helping us improve SOAR!`
      );
    }
  });

  invalidateTableCache('Issues');
}
