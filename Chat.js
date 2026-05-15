const SOAR_SYSTEM_PROMPT = `
You are the SOAR AI Assistant. SOAR is a lightweight, Google-native project management system built on Google Apps Script and Google Sheets.
Your job is to answer user questions about how to use SOAR. Keep your answers concise, friendly, and helpful.

`;

function askGeminiAssistant(conversationHistory, userContext) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return JSON.stringify({ success: false, error: 'AI Assistant is not configured (Missing API Key).' });
  }

  // Load the Tutorial HTML to inject into the system prompt
  let tutorialContent = "";
  try {
    tutorialContent = HtmlService.createHtmlOutputFromFile('Tutorial').getContent();
  } catch(e) {
    tutorialContent = "Tutorial file not found.";
  }

  const DYNAMIC_SYSTEM_PROMPT = `
${SOAR_SYSTEM_PROMPT}

=== SOAR USER TUTORIAL & PLAYBOOK ===
${tutorialContent}

=== CURRENT USER STATE ===
${userContext}
  `;

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const contents = conversationHistory.map(msg => ({
    role: msg.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: msg.content }]
  }));

  const payload = {
    system_instruction: { parts: [{ text: DYNAMIC_SYSTEM_PROMPT }] },
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
    return JSON.stringify({ 
      success: true, 
      text: aiResponseText,
      debugPayload: payload
    });

  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to fetch AI response.' });
  }
}