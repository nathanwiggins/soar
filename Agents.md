# AI Agent Instructions for SOAR

Welcome! You are an AI agent assisting with **SOAR**, a lightweight, Google-native project management system built on Google Apps Script and Google Sheets. When contributing to this repository, you must strictly adhere to the following rules:

## 1. Documentation is the Source of Truth
Every time you make a change that affects the user interface, feature set, data model, or API, you **MUST** update `README.md`. 
- Keep the `Spreadsheet Setup` and `Data Dictionary` sections perfectly aligned with the backend code.
- Ensure all API references match the actual Apps Script function signatures and return types.
- Ensure that Tutorial.html exactly matches README.md

## 2. The UI Playbook Rule
The section titled `## SOAR Tutorial and UI Playbook` in `README.md` serves as the live knowledge base for the in-app SOAR AI Assistant. 
- Do not invent or hallucinate UI elements. Always use exact button names, tab names, and modal titles currently present in `App.js.html` and `Index.html`.
- If you change a UI element (e.g., renaming a button), you MUST update the `UI Playbook` section in `README.md` to reflect this change.
- **DO NOT manually edit `Tutorial.html`.** It is auto-generated from `README.md` via a build script.

## 3. Tech Stack & Constraints
- **Backend:** Google Apps Script (V8). Use `google.script.run` for client-server communication. Backend functions should return JSON strings for easy frontend parsing.
- **Frontend:** Vue 3 (via CDN), Tailwind CSS (via CDN). Single Page Application structure split between `App.js.html` and `Index.html`.
- **Database:** Google Sheets. Do not introduce external databases or dependencies.

## 4. Cache & Data Versioning
- Any time you write to the Google Sheet (create, update, delete rows), you must call `invalidateTableCache('TableName')`. This bumps the global data version and triggers frontend Vue components to reactively fetch the fresh data.

## 5. Error Handling
- Never wrap `google.script.run` calls in silent try/catch blocks on the frontend without alerting the user. 
- Backend mutation functions should always catch errors and return `{ success: false, error: "Human readable message" }`.