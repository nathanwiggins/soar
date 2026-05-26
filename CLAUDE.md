# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## AI Agent Instructions for SOAR

Welcome! You are an AI agent assisting with **SOAR**, a lightweight, Google-native project management system built on Google Apps Script and Google Sheets.

Before starting work on new requests, always perform a `git pull` to ensure that the local codebase is up to date.

When contributing to this repository, every time you make a change that affects the user interface, feature set, backend code, data model, or API, you **MUST** update `README.md` AND `Tutorial.html`.
- Ensure that `README.md` sections are perfectly aligned with the backend code.
- Ensure that `Tutorial.html` receives any updates related to the user experience so that the AI Assistant can answer the user's questions accurately.
- Ensure that common sections between `README.md` and `Tutorial.html` are synced with each other.

Do not include code comments. This codebase should be free of all code comments.

After making changes, be sure to do `clasp push` to send the changes to the dev environment so the user can view them.

---

## PR Review Process

When asked to review a pull request:

1. **Keep it brief** — summarize what the PR does in plain language and flag only critical issues. Skip minor style notes unless asked for a deeper review.
2. **Always check for missing docs** — every PR that touches the UI, features, backend, or data model must update `README.md` and `Tutorial.html`. Flag it if either is missing.
3. **Before merging** — fix any flagged issues on the PR branch, then push the fix to the branch before merging.
4. **Merging with conflicts** — merge locally (`git merge origin/<branch>`), resolve conflicts manually, commit, then `clasp push` before `git push origin main`. Walk the user through each conflict and explain the reasoning before resolving.
5. **After merging** — commit any post-merge fixes, `clasp push`, then `git push origin main`.

---

## Development Commands

```bash
clasp push          # Deploy all local files to the Apps Script dev environment
clasp pull          # Pull current Apps Script project files locally
clasp open          # Open the Apps Script editor in a browser
clasp logs          # Stream Stackdriver logs from the live deployment
```

There are no local build steps, test runners, or linters — the runtime is Google Apps Script (V8 engine).

---

## Architecture Overview

SOAR is a **serverless Google Apps Script web app** with no custom server or external database. All server-side code runs in Google's Apps Script runtime; all data lives in a Google Sheets spreadsheet.

### Request / Response Model

Apps Script web apps do not support REST — the client communicates exclusively via `google.script.run`:

```
Browser (Vue 3 SPA)  →  google.script.run.<functionName>(args)  →  Apps Script server
                     ←  .withSuccessHandler(cb) / .withFailureHandler(cb)
```

Every public server function that the client calls must be a top-level `function` declaration in one of the `.js` files. Functions return JSON strings (not objects); the client parses them. There is no `doPost`.

### Initial Data Load & Caching

1. On mount, the client checks `localStorage` for `soar_initial_payload_v4`.
2. If cached, it immediately calls `getGlobalVersionHash()` to compare hashes.
3. If the hash matches, it hydrates from cache (fast). If it differs, it calls `getInitialPayload()` to fetch all data in one round-trip, then writes the new payload back to `localStorage`.
4. `getInitialPayload()` (in `Bootstrap.js`) fetches every sheet table in a single execution and returns one large JSON object containing `users`, `projects`, `tasks`, `subtasks`, `assignments`, `agendas`, `agendaShares`, `comments`, `currentUserSettings`, and `versionHash`.

The version hash (`buildGlobalVersionHash`) combines the spreadsheet's Drive last-modified timestamp with an app-level `soar_data_version` script property that is bumped on every write.

### Server-Side Caching Layers

| Layer | Scope | TTL | Purpose |
|---|---|---|---|
| `REQUEST_CACHE` (in-memory object) | Single execution | Execution lifetime | Deduplicate sheet reads within one request |
| `CacheService.getScriptCache()` | All executions | 5 minutes | Cache the Users table across requests |
| `PropertiesService` | Persistent | Until bumped | Store ID counters and `soar_data_version` |

`invalidateTableCache(sheetName)` clears both `REQUEST_CACHE` and (for Users) the script cache, then bumps the data version. Call it after every write.

### Data Model

Eight Google Sheets tabs act as tables. Column headers are the field names — they must match exactly:

- **Users** — `User_ID`, `Name`, `Email`, `Manager_ID`, `Profile_Pic_Url`
- **Projects** — `Project_ID`, `Project_Title`, `Description`, `Status`, `Due_Date`, `Creator_ID`, `Color_Scheme`, `Sort_Order`
- **Tasks** — `Task_ID`, `Project_ID`, `Task_Title`, `Description`, `Status`, `Priority`, `Due_Date`, `Creator_ID`, `Sort_Order`
- **Subtasks** — `Subtask_ID`, `Task_ID`, `Subtask_Title`, `Status`, `Sort_Order`
- **Assignments** — `Assignment_ID`, `Assignee_ID`
- **Comments** — `Comment_ID`, `Topic_ID`, `Topic_Type`, `Commenter_ID`, `Content`, `Timestamp`, `Is_Resolved`
- **Agendas** — `Agenda_ID`, `Title`, `Description`, `Agenda_Date`, `Content_JSON`, `Creator_ID`, `Created_Date`
- **Sharing** — `Agenda_ID`, `User_ID`

IDs are generated via `generateNextId(sheetName, prefix)` (e.g., `TASK-00000042`), using a script-property counter with a script lock to prevent races.

### Frontend

`Index.html` is the HTML shell. It loads Tailwind CSS, Vue 3, SortableJS, vuedraggable, and marked from CDNs, then uses Apps Script's `<?= include('App.js.html') ?>` templating to inline the Vue app script.

`App.js.html` is a single-file Vue 3 Composition API app (~3 200 lines). All state lives in `setup()` refs. Key patterns:
- `memoizeBySignature(cacheKey, signature, factory)` — client-side memoization for expensive computed values; invalidated when the signature (usually a data hash) changes.
- `google.script.run` calls always use `.withSuccessHandler` / `.withFailureHandler`; responses are `JSON.parse`d.
- Dark mode, font scale, and notification preferences are persisted via `persistCurrentUserSettings` to `PropertiesService`.

### Backend Module Responsibilities

| File | Responsibility |
|---|---|
| `Code.js` | `doGet` entry point and `include` helper |
| `Bootstrap.js` | `getInitialPayload`, `getGlobalVersionHash`, user identity & profile photo sync |
| `DataStore.js` | All sheet reads/writes, ID generation, cache, version hashing |
| `Utilities.js` | Shared helpers: email normalization, date parsing, `safeSendEmail`, header indexing |
| `Projects.js` | Project CRUD, reordering, color scheme |
| `Tasks.js` | Task CRUD, status/priority normalization, assignment validation |
| `Subtasks.js` | Subtask CRUD and reordering |
| `Comments.js` | Comment CRUD, mention extraction, resolve |
| `Users.js` | User creation, profile updates, assignable-user scoping |
| `Settings.js` | Per-user notification settings via `PropertiesService` |
| `Notifications.js` | Email notification dispatch on task/comment events |
| `Agendas.js` | Agenda CRUD and sharing |
| `Chat.js` | Gemini AI assistant integration (requires `GEMINI_API_KEY` in Script Properties) |

### Adding a New Data Entity

1. Add a new sheet tab with exact column headers.
2. Add CRUD functions in a new `.js` file following the pattern in existing modules (use `getTableData`, `invalidateTableCache`, `generateNextId`, `appendRows`, etc.).
3. Include the new table in `getInitialPayload` in `Bootstrap.js`.
4. Update `README.md` spreadsheet setup section and `Tutorial.html`.
