function generateIssueEmailParts(issueSummary) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { shortDescription: 'Your reported issue', description: issueSummary };

  const prompt = `Given the following SOAR support issue summary, return a JSON object with two fields:
1. "shortDescription": a concise title (max 7 words) suitable for an email subject line. (e.g. Fix missing fields, Add new feature, etc.)
2. "description": a short phrase (not a full sentence) that fits naturally into the blank here: "The issue you reported regarding ___ has been resolved." It should read as a natural continuation of that sentence — for example: "task statuses not updating correctly" or "the dashboard failing to load" or "adding a new color scheme".

Issue summary:
${issueSummary}

Respond with only valid JSON, no markdown or code fences.`;

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`;
  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }] }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(response.getContentText());
    const text = json.candidates[0].content.parts[0].text.trim();
    const parsed = JSON.parse(text);
    return {
      shortDescription: parsed.shortDescription || 'Your reported issue',
      description: parsed.description || issueSummary
    };
  } catch (e) {
    Logger.log(`generateIssueEmailParts: ${e.message}`);
    return { shortDescription: 'Your reported issue', description: issueSummary };
  }
}

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
  const scriptProperties = PropertiesService.getScriptProperties();
  const scriptPropertiesCache = scriptProperties.getProperties();
  const githubPat = scriptPropertiesCache['GITHUB_PAT'];
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

    const responseCode = githubResponse.getResponseCode();

    if (responseCode === 410) {
      const newRow = [...row];
      newRow[headerIndex.Status] = 'Deleted';
      updateRowValues(sheet, sheetRowIndex, newRow);
      return;
    }

    if (responseCode !== 200) {
      Logger.log(`syncDailyGitHubStatus: GitHub API error for issue #${issueNumber}: ${githubResponse.getContentText()}`);
      return;
    }

    const githubIssue = JSON.parse(githubResponse.getContentText());
    if (githubIssue.state === 'closed') {
      const newRow = [...row];
      newRow[headerIndex.Status] = 'Complete';
      updateRowValues(sheet, sheetRowIndex, newRow);

      if (isNotificationEnabledForEmail(row[headerIndex.User_Email], 'ticketFollowUp', scriptPropertiesCache)) {
        const { shortDescription, description } = generateIssueEmailParts(row[headerIndex.Issue_Description]);
        safeSendEmail(
          row[headerIndex.User_Email],
          `[SOAR Issue] ${shortDescription}`,
          `Hello,\n\nThe issue you reported regarding ${description} has been resolved by our development team.\n\nThank you for helping us improve SOAR!`
        );
      }
    }
  });

  invalidateTableCache('Issues');
}
