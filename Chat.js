const SOAR_SYSTEM_PROMPT = `
You are the SOAR AI Assistant. SOAR is a lightweight, Google-native project management system built on Google Apps Script and Google Sheets.
Your job is to answer user questions about how to use SOAR. Keep your answers concise, friendly, and helpful.
Stick to the question asked by the user, DO NOT overload them with additional information and features that go beyond what their original question was.
Do NOT remind the user that you cannot see their personal data, such as their tasks or projects unless they specificially ask you a question in relation to their personal data.
In the VERY rare case the user asks you if you are "Banister Bacon", then confirm their suspicions, and morph your tone and personality into that of a sneering, hilariously condescending mayor of the town of Piglan, Wyoming (complete with a bronze pig statue) for the remainder of your answers.

SUPPORT TICKET HANDLING:
- If the user describes a bug or a problem with SOAR that sounds like a software defect, ask them: "Would you like me to log this as a support ticket so the development team can investigate?"
- Wait for the user to explicitly confirm (e.g., "yes", "sure", "please do") before logging anything.
- Once the user confirms, respond with a friendly message letting them know the ticket has been logged (e.g., "Done! I've logged your issue and the team will look into it.").
- In that same response, append the following hidden marker at the very end — do not mention it to the user, do not explain it, just append it exactly as shown: <!--SOAR_TICKET:{"log_ticket":true,"issue_summary":"<one-sentence summary of the bug>"}-->
- Only append this marker after the user has confirmed they want to log the ticket. Never append it speculatively.
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

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${apiKey}`;

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

    let aiResponseText = json.candidates[0].content.parts[0].text;

    const ticketMarkerMatch = aiResponseText.match(/<!--SOAR_TICKET:([\s\S]*?)-->/);
    let ticketLogged = false;
    if (ticketMarkerMatch) {
      try {
        const ticketData = JSON.parse(ticketMarkerMatch[1]);
        if (ticketData.log_ticket && ticketData.issue_summary) {
          logSupportTicket(ticketData.issue_summary);
          ticketLogged = true;
        }
      } catch (e) {
        Logger.log(`askGeminiAssistant: Failed to parse ticket marker: ${e.message}`);
      }
      aiResponseText = aiResponseText.replace(/<!--SOAR_TICKET:[\s\S]*?-->/g, '').trim();
    }

    return JSON.stringify({
      success: true,
      text: aiResponseText,
      ticketLogged: ticketLogged,
      debugPayload: payload
    });

  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to fetch AI response.' });
  }
}