function logSupportTicket(issueSummary) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
  if (!sheet) {
    return JSON.stringify({ success: false, error: 'Issues sheet not found.' });
  }

  const issueId = generateNextId('Issues', 'ISSUE');
  const timestamp = new Date().toISOString();
  const userEmail = normalizeEmail(getCurrentUser());

  appendRows(sheet, [[issueId, timestamp, userEmail, issueSummary, 'New', '']]);
  invalidateTableCache('Issues');

  return JSON.stringify({ success: true, issueId: issueId });
}

function processWeeklyTickets() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const githubPat = PropertiesService.getScriptProperties().getProperty('GITHUB_PAT');

  if (!apiKey) {
    Logger.log('processWeeklyTickets: Missing GEMINI_API_KEY');
    return;
  }
  if (!githubPat) {
    Logger.log('processWeeklyTickets: Missing GITHUB_PAT');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const headerIndex = getHeaderIndex(headers);

  const newIssueRows = data.slice(1)
    .map((row, i) => ({ row, sheetRowIndex: i + 2 }))
    .filter(({ row }) => row[headerIndex.Status] === 'New');

  if (newIssueRows.length === 0) return;

  const issueList = newIssueRows.map(({ row }, i) =>
    `${i + 1}. Issue ID: ${row[headerIndex.Issue_ID]}, Description: "${row[headerIndex.Issue_Description]}"`
  ).join('\n');

  const prompt = `You are a software quality analyst. Review these reported issues and classify each as either "bug" (a real defect in the software) or "user_error" (user misunderstanding or incorrect usage). Respond ONLY with a JSON array like: [{"issue_id": "ISSUE-00000001", "classification": "bug"}, ...]\n\nIssues:\n${issueList}`;

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;
  const geminiResponse = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  });

  if (geminiResponse.getResponseCode() !== 200) {
    Logger.log(`processWeeklyTickets: Gemini API error: ${geminiResponse.getContentText()}`);
    return;
  }

  let classifications;
  try {
    const json = JSON.parse(geminiResponse.getContentText());
    const raw = json.candidates[0].content.parts[0].text;
    const jsonMatch = raw.match(/\[[\s\S]*\]/);
    classifications = JSON.parse(jsonMatch[0]);
  } catch (e) {
    Logger.log(`processWeeklyTickets: Failed to parse Gemini response: ${e.message}`);
    return;
  }

  classifications.forEach(({ issue_id, classification }) => {
    const rowEntry = newIssueRows.find(({ row }) => row[headerIndex.Issue_ID] === issue_id);
    if (!rowEntry) return;

    const { row, sheetRowIndex } = rowEntry;
    const newRow = [...row];

    if (classification === 'bug') {
      const githubResponse = UrlFetchApp.fetch('https://api.github.com/repos/nathanwiggins/soar/issues', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: `Bearer ${githubPat}`, 'User-Agent': 'SOAR-App' },
        payload: JSON.stringify({
          title: `[User Report] ${row[headerIndex.Issue_Description].toString().substring(0, 100)}`,
          body: `**Reported by:** ${row[headerIndex.User_Email]}\n**Timestamp:** ${row[headerIndex.Timestamp]}\n\n${row[headerIndex.Issue_Description]}`
        }),
        muteHttpExceptions: true
      });

      if (githubResponse.getResponseCode() === 201) {
        const githubIssue = JSON.parse(githubResponse.getContentText());
        newRow[headerIndex.Status] = 'Pending';
        newRow[headerIndex.GitHub_Issue_Number] = githubIssue.number;
        updateRowValues(sheet, sheetRowIndex, newRow);
      } else {
        Logger.log(`processWeeklyTickets: GitHub API error for ${issue_id}: ${githubResponse.getContentText()}`);
      }
    } else {
      newRow[headerIndex.Status] = 'User Error';
      updateRowValues(sheet, sheetRowIndex, newRow);
    }
  });

  invalidateTableCache('Issues');
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
