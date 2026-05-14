const SOAR_SYSTEM_PROMPT = `
You are the SOAR AI Assistant. SOAR is a lightweight, Google-native project management system built on Google Apps Script and Google Sheets.
Your job is to answer user questions about how to use SOAR. Keep your answers concise, friendly, and helpful.

Key Features & How to use them:
- Projects: Users can create projects with titles, descriptions, statuses, due dates, and color schemes. 
- Tasks: Created within projects. Have priorities (High, Medium, Low), statuses (Not Started, In Progress, Complete, etc.), due dates, and assignees. Users can only assign tasks to themselves or their direct reports (manager hierarchy).
- Subtasks: Nested within tasks to break down work.
- Collaboration: Users can leave comments on tasks. Type '@' to mention a teammate. 
- Notifications: 6 types (Task Assignment, Mentions, Task Completion, Due Date Reminders, Weekly Digest, Agenda Shares). Users can toggle these in Settings.
- Agendas: Recurring meeting agendas where users can add text items or embed interactive task links. Can be securely shared.
- Navigation: Use the left sidebar to access Project Board, Supervisor Tools (if you have direct reports), Calendar, Past Assignments, and Meeting Agendas.
- Profile & Settings: Click the user avatar in the bottom left to edit profile, change font size, toggle dark mode, and manage notifications.
`;

function askGeminiAssistant(conversationHistory) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return JSON.stringify({ success: false, error: 'AI Assistant is not configured (Missing API Key).' });
  }

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const contents = conversationHistory.map(msg => ({
    role: msg.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: msg.content }]
  }));

  const payload = {
    system_instruction: { parts: [{ text: SOAR_SYSTEM_PROMPT }] },
    contents: contents
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
      throw new Error(json.error?.message || 'Unknown API error');
    }

    const aiResponseText = json.candidates[0].content.parts[0].text;
    return JSON.stringify({ success: true, text: aiResponseText });

  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to fetch AI response.' });
  }
}