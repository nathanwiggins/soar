# Soar — Project Management for Google Workspace

**Soar** is a lightweight, Google-native project management system built on **Google Apps Script** and **Google Sheets**. It provides project tracking, task management, subtasks, task comments with mentions, meeting agendas with sharing, team-supervisor views, calendar views, configurable email notifications, and an optional in-app SOAR AI Assistant—all without a custom server or external database.

**Perfect for**: Small to mid-sized teams already using Google Workspace who want project management without complex setup or external infrastructure.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [SOAR Tutorial and UI Playbook](#soar-tutorial-and-ui-playbook)
3. [Features](#features)
4. [Architecture](#architecture)
5. [Data Model](#data-model)
6. [Configuration](#configuration)
7. [API Reference](#api-reference)
8. [Development](#development)
9. [Troubleshooting](#troubleshooting)

---

## Quick Start

### Prerequisites

- Google account with Apps Script access (part of Google Workspace)
- Permission to create a new Google Sheet
- Permission to deploy Apps Script web apps
- Optional, only for the in-app AI Assistant: a Gemini API key saved in Script Properties as `GEMINI_API_KEY`

### Deployment (5 minutes)

1. **Create a new Google Sheet**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Click **New** → **Blank spreadsheet**
   - Name it `Soar` (or your preferred name)
   - Keep this spreadsheet as the active spreadsheet for the Apps Script project.

2. **Create the data structure**
   - Create 9 sheet tabs with these exact names (right-click a sheet tab → **Insert sheet**):
     - `Users`
     - `Projects`
     - `Tasks`
     - `Subtasks`
     - `Comments`
     - `Assignments`
     - `Agendas`
     - `Sharing`
     - `Sessions`
   - Add the exact header rows shown in [Spreadsheet Setup](#spreadsheet-setup). Column names must match exactly.

3. **Create Apps Script project**
   - In your Google Sheet, go to **Extensions** → **Apps Script**.
   - A new Apps Script project will open.
   - Delete any default content in `Code.gs`.

4. **Add the source code**
   - Create Apps Script script files (`.gs`) for each `.js` file in this repository, and paste the content into them:
     - [Utilities.js](Utilities.js)
     - [DataStore.js](DataStore.js)
     - [Users.js](Users.js)
     - [Projects.js](Projects.js)
     - [Tasks.js](Tasks.js)
     - [Subtasks.js](Subtasks.js)
     - [Comments.js](Comments.js)
     - [Agendas.js](Agendas.js)
     - [Notifications.js](Notifications.js)
     - [Settings.js](Settings.js)
     - [Bootstrap.js](Bootstrap.js)
     - [Chat.js](Chat.js)
     - [Code.js](Code.js) — paste this into the existing `Code.gs` file if you did not create a separate `Code` script file.
   - Create Apps Script HTML files for:
     - [Index.html](Index.html)
     - [App.js.html](App.js.html)
     - [Tutorial.html](Tutorial.html)
   - Update `appsscript.json` with the manifest in [Configuration](#configuration).

5. **Optional: configure the SOAR AI Assistant**
   - In Apps Script, open **Project Settings** → **Script Properties**.
   - Add a property named `GEMINI_API_KEY` with your Gemini API key.
   - If this property is missing, the red chat bubble still appears, but assistant requests return: `AI Assistant is not configured (Missing API Key).`

6. **Deploy as web app**
   - Click **Deploy** → **New deployment**.
   - Type: select **Web app**.
   - Execute as: **User accessing the web app** (matches `executeAs: USER_ACCESSING` in the manifest).
   - Who has access: **Anyone** (matches `access: ANYONE` in the manifest).
   - Click **Deploy** and authorize requested scopes.
   - Open the deployment URL.

7. **Initialize your account**
   - On first load, SOAR checks whether the signed-in Google email exists in the `Users` tab.
   - If not, the **Create Account** modal appears.
   - Your **Email** field is prefilled and disabled.
   - Enter **Name** and optionally choose **Manager (Optional)** from existing users.
   - Click **Create Account**.

### First Steps in the App

- **Create a project**: On **Project Board**, click **New Project**, fill **Project Title**, optional **Due Date**, **Status**, **Color Scheme**, and optional **Description**, then click **Create Project**.
- **Create a task**: Inside a project column, click **+ Add Task** (not “New Task”), fill the **Add Task** modal, select at least one user under **Assigned To**, then click **Create Task**.
- **Update a task quickly**: Use the status pill/dropdown on a task card to choose `Not Started`, `Upcoming`, `Review`, `In Progress`, `Ongoing`, `On Hold`, `Cancelled`, or `Complete`.
- **Open details**: Click a project title to open **Project Details**. Click a task card to open **Task Details**.
- **Comment on tasks**: Click the speech-bubble icon on a task card to open **Comments**, type in **Write a comment...**, optionally use `@` mentions, then click **Post Comment**.
- **Create an agenda**: Go to **Meeting Agendas**, click **New Agenda**. The agenda editor opens with no sessions. Click **New Session** and choose blank or copy-from-previous, add headers/items/tasks, optionally click **Share**, then click **Save Session**.
- **Customize**: Open the user menu at the lower-left, then use **Profile**, **Settings**, or **Dark Mode**.

---

## SOAR Tutorial and UI Playbook

This section is written as a practical, non-technical guide. It uses the exact in-app names for tabs, buttons, fields, modals, and actions so it can be safely used as context for an AI Assistant.

### Main Layout

SOAR has a left sidebar, a top header, a main work area, a lower-left user menu, and a lower-right chat bubble.

#### Left sidebar tabs

- **Project Board**: The default work board. Shows project columns and task cards.
- **Supervisor Tools**: Only available to users who have direct reports. Shows selected direct reports' assigned work.
- **Calendar**: Month view of project due dates and task due dates.
- **Past Assignments**: Completed tasks assigned to the current user.
- **Meeting Agendas**: Agenda cards under **My Agendas** and **Shared With Me**.

#### Top header actions

- On **Project Board**, the primary red button is **New Project**.
- On **Meeting Agendas**, the primary red button is **New Agenda**.
- On **Supervisor Tools**, the header has a team-member selector whose default text is **Select team members**.
- On **Calendar**, the header has previous-month and next-month arrow buttons, a **Today** button, and the current month label.

#### Lower-left user menu

Click your name/avatar at the bottom of the sidebar to open:

- **Profile**: Opens **My Profile**.
- **Settings**: Opens **Settings**.
- **Dark Mode**: Toggles dark mode. The menu displays `On` or `Off`.

#### Lower-right SOAR Assistant

- Click the red circular message button to open **SOAR Assistant**.
- Type in the **Ask a question...** box and submit with the paper-plane button.
- The chat window can be repositioned by dragging its header bar.
- The assistant sends up to 10 recent messages of history plus a small current-page context (active tab and whether the user has direct reports) to `askGeminiAssistant()`.
- Assistant responses are rendered as formatted markdown.
- A privacy disclaimer is displayed at the bottom of the chat: messages are processed by the Gemini API, data may be used to train Google's AI models, and users should not share private or sensitive information.
- The assistant needs Script Property `GEMINI_API_KEY`; otherwise, it responds with a configuration error.

### Onboarding: Create Account

When the signed-in Google user is not in the `Users` sheet, SOAR opens the **Create Account** modal.

Fields and controls:

- **Email**: Prefilled from the signed-in Google account and disabled.
- **Name**: Required; placeholder is **Enter your full name**.
- **Manager (Optional)**: Dropdown; default option is **No manager selected**.
- **Create Account**: Creates the user.

Important behavior:

- The manager dropdown only contains users already in SOAR.
- The app can automatically sync a Google profile photo URL into `Profile_Pic_Url` when possible.
- If the selected manager has notifications enabled, SOAR can notify them that the account was created.

### Project Board: Everyday Task and Project Work

The **Project Board** displays each visible project as a column. Within each column, task cards are sorted by priority and due date.

#### Visibility on Project Board

By default, a user sees:

- projects directly assigned to them, and
- projects containing open tasks assigned to them.

Completed tasks are not shown on the main **Project Board** after they are complete; they appear in **Past Assignments** for assigned users. Project visibility can expand when the app's internal `showAllWorkItems` state is enabled, but there is no visible button in the current UI for end users to toggle that state.

#### Project columns

Each project column shows:

- a colored dot using the project's **Color Scheme**;
- the project title as a clickable button that opens **Project Details**;
- the project due date (when set), color-coded: yellow if due within 7 days, orange if due within 1 day, red if overdue;
- the number of visible tasks in that project;
- draggable task cards;
- a dashed **+ Add Task** button at the bottom.

Projects can be reordered by dragging the project header area. Only the project creator can drag their own project column. The new order is saved per-user via `saveUserSortOrder()` and does not affect the order other users see.

#### Task cards

Each task card shows:

- status dropdown/pill with title **Update task status**;
- subtask count indicator (list-check icon and number) next to the status pill, shown only when the task has subtasks;
- speech-bubble comments button with the number of unresolved comments;
- task title;
- assignee avatars or initials;
- priority icon/label when priority is set;
- due date when set, color-coded: yellow if due within 7 days, orange if due within 1 day, red if overdue (gray for completed tasks).

Task cards can be dragged between project columns. Moving a task to a different project requires the user to be the creator of both the source and target project. When dragging over an unauthorized project, a red "Not authorized to move here" banner appears; authorized targets show a blue "Move to: [Project]" banner. Cross-project moves persist the task's new `Project_ID` via `moveTaskToProject()`, and the user's task order is saved per-user via `saveUserSortOrder()`. Press **Cmd/Ctrl+Z** to undo the last cross-project move.

### Creating a Project

1. Go to **Project Board**.
2. Click **New Project** in the top header.
3. The **New Project** modal opens.
4. Complete fields:
   - **Project Title**: required; placeholder **Enter project title**.
   - **Due Date**: optional date picker.
   - **Status**: `Not Started`, `In Progress`, `Completed`, or `Delayed`.
   - **Color Scheme**: `SUU Red (Default)`, `Sunset Orange`, `Amber Gold`, `Emerald Green`, `Ocean Teal`, `Sky Blue`, `Deep Indigo`, `Soft Violet`, `Rose Pink`, `Slate Gray`, or `Pearl White`.
   - **Description**: optional; placeholder **Describe the project**.
5. Click **Create Project**. To exit without saving, click **Cancel** or the **X** icon.

What SOAR records:

- `Project_ID` generated as `P-00000001`, etc.
- `Project_Title`, `Description`, `Status`, `Created_Date`, `Due_Date`, `Creator_ID`, and `Color_Scheme`.
- An assignment row assigning the project to the creator.

### Editing or Deleting a Project

1. On **Project Board**, click a project title.
2. The **Project Details** modal opens.
3. Click **Edit Project**.
4. Edit fields:
   - **Project Title**
   - **Date Due**
   - **Status**
   - **Color Scheme**
   - **Created By** (display-only)
   - **Description**
5. Click **Save Changes**.

Other buttons:

- **Delete Project**: Deletes the project row, tasks in that project, and assignment rows for those deleted tasks. It does not currently remove the project creator assignment row from `Assignments`.
- **Cancel**: Cancels edit mode.
- **Close**: Closes the modal when not editing.

### Creating a Task

1. On **Project Board**, find the target project column.
2. Click **+ Add Task** at the bottom of that project column.
3. The **Add Task** modal opens and shows `Project: {Project_Title}` below the heading.
4. Complete fields:
   - **Task Title**: required; placeholder **Enter task title**.
   - **Due Date**: optional date picker.
   - **Priority**: optional button selection: `High`, `Medium`, or `Low`.
   - **Assigned To**: required by backend validation; opens a checkbox dropdown of assignable users.
   - **Description**: optional; placeholder **Describe the task**.
   - **Subtasks**: optional; type into **Type a subtask and press enter...** and press Enter or click **Add**.
5. Click **Create Task**. To exit without saving, click **Cancel** or the **X** icon.

Important behavior:

- New tasks always start with status `Not Started`.
- The current user may assign tasks only to themselves and users in their reporting tree (direct and indirect reports). Existing assignees can remain during edits even if they are outside the current assignable set.
- At least one assignee is required when creating or updating a task.
- Creating a task can send **Task assignments** notifications to selected assignees, depending on each recipient's settings.

### Task Details: Editing, Completing, Deleting

Click a task card to open **Task Details**.

Controls and fields:

- **Edit Task**: enables editing.
- **Task Name**: task title.
- **Description**
- **Subtasks**: checkboxes, editable titles in edit mode, drag handles in edit mode, and delete controls.
- **Task Status**: `Not Started`, `Upcoming`, `Review`, `In Progress`, `Ongoing`, `On Hold`, `Cancelled`, or `Complete`.
- **Associated Project**: project dropdown available while editing.
- **Priority**: dropdown with `None`, `High`, `Medium`, `Low`.
- **Date Created**: display-only.
- **Date Due**: editable date picker in edit mode.
- **Completed By** and **Completed At**: populated when the task is completed.
- **Assigned To**: assignee dropdown plus selected assignee rows.

Footer buttons:

- **Complete Task**: sets the task status to `Complete`, records `Completed_By`, and records `Completed_At`.
- **Complete**: displayed in the same button position when the task is already complete.
- **Delete Task**: deletes the task and related assignments.
- **Save Changes**: saves edits.
- **Cancel**: cancels edit mode.
- **Close**: closes the modal when not editing.

Completed-task behavior:

- Completed tasks assigned to you appear in **Past Assignments**.
- Completed tasks with due dates in the past can be purged by `purgeCompletedTasksPastDue()`.
- Managers can receive task-completion notifications for work completed by reports.

### Subtasks

Subtasks belong to tasks and have their own IDs (`S-00000001`, etc.).

In **Add Task**:

- Use the **Subtasks** input with placeholder **Type a subtask and press enter...**.
- Press Enter or click **Add** to stage a subtask before clicking **Create Task**.

In **Task Details**:

- Click **Edit Task** to edit subtask titles, add new subtasks, delete subtasks, or drag to reorder subtasks.
- A subtask checkbox toggles status between `Incomplete` and `Complete`.

### Comments and Mentions

Comments are currently task comments. Despite legacy data-model support for `Topic_Type`, the active UI/backend flow validates comments against tasks and writes `Topic_Type = Task`.

How to comment:

1. On a task card, click the speech-bubble comments icon.
2. The **Comments** modal opens.
3. Type in **Write a comment...**.
4. Type `@` to show mention suggestions. Suggestions use handles derived from each user's email local-part, or a sanitized display-name fallback.
5. Click **Post Comment**.

Comment controls:

- **Resolve**: marks the comment resolved; resolved comments no longer appear in the active comments list.
- **Delete**: permanently deletes the comment row.

Mention behavior:

- Mentions are recognized with handles like `@first.last` or `@jane`, not display names with spaces.
- Mention notifications use the **Comments and mentions** notification preference.

### Supervisor Tools

**Supervisor Tools** is visible only if the current user has direct reports in the `Users` sheet (`Manager_ID` points to the current user's `User_ID`).

How it works:

1. Click **Supervisor Tools** in the sidebar.
2. Use the header selector, default text **Select team members**.
3. Check one or more direct reports.
4. SOAR shows open tasks assigned to selected users and projects containing those tasks.

Notes:

- The selector lists direct reports, not the full indirect reporting tree.
- The board itself can still show tasks in project columns, and the **+ Add Task** button remains visible.

### Calendar

The **Calendar** tab is a month view.

Header controls:

- left arrow: previous month;
- **Today**: return to the current month;
- month label: current displayed month and year;
- right arrow: next month.

Calendar entries:

- Project due dates appear as **Project Due:** entries with a folder icon.
- Task due dates appear as task-title entries.
- Drag and drop both task entries and project-due entries onto another day cell to immediately update their due dates. Only the project creator can drag their own tasks and projects.
- While dragging, the hovered day cell shows a "Move [type] to [date]" tooltip so you can confirm the target before dropping.
- Press **Cmd/Ctrl+Z** to undo the last calendar date change.
- Click a project-due entry to open **Project Details**.
- Click a task entry to open **Task Details**.

### Past Assignments

The **Past Assignments** tab shows completed tasks assigned to the current user.

- If there are none, SOAR displays **No past assignments yet.** and **Completed tasks will appear here automatically.**
- Each completed task card shows task title, project title, a `Completed` badge, due date, and **Delete Permanently**.
- Click a card to open **Task Details**.
- Click **Delete Permanently** to delete the completed task and related assignments.

### Meeting Agendas

The **Meeting Agendas** tab has two sections:

- **My Agendas**: agendas created by the current user.
- **Shared With Me**: agendas shared with the current user.

Agenda cards display:

- Section count badge (from the latest session).
- Title.
- Description (if set).
- Creator and shared-user avatars.
- Latest session date (or "No sessions yet" if no sessions exist).

Each agenda is a recurring **template**. Every time the team meets, the agenda creator adds a new **session** to the agenda. Sessions are date-stamped meeting instances that hold the actual agenda content (headers, items, linked tasks). Past sessions are read-only and browsable via the **← Older** / **Newer →** navigation bar at the top of the editor.

Creating an agenda:

1. Go to **Meeting Agendas**.
2. Click **New Agenda**.
3. SOAR creates the agenda and opens the editor. No sessions exist yet.
4. Edit the title in the title input at the top of the modal.
5. Add a description in the description textarea.
6. Click **New Session** (visible to the agenda creator) and choose **Blank** or **Copy from previous session**.
7. Set the date shown in the session navigation bar.
8. Click **+ Add Header** to create a section.
9. Inside a section:
   - click **+ Text Item** to add a free-text agenda item;
   - use **+ Link Task** dropdown to embed a task from a visible project;
   - linked task cards display the task title, status badge, assignee avatars, priority, and due date;
   - click **View** on a linked task (visible on hover) to open **Task Details**.
10. Drag agenda items to reorder them within sections.
11. Click **Save Session**.

Navigating sessions:

- The session navigation bar shows **Session N of M · [date]**.
- Click **← Older** to browse earlier sessions (right → older).
- Click **Newer →** to return to more recent sessions.
- Only the **latest** session (index 0) is editable. All past sessions are read-only.

Managing sessions:

- **New Session** button (owner only): creates a new session, either blank or pre-filled from the previous session's content. The new session date defaults to today.
- **Delete Session** button (owner only): deletes the currently viewed session. Only visible when more than one session exists.
- Deleting an agenda deletes all its sessions.

Sharing an agenda:

1. In the agenda editor, click **Share**.
2. The **Share Agenda** pop-up opens.
3. Use **Add people (Type name or email)...** to search users.
4. Click a suggestion to add access.
5. The creator appears as `{Name} (You)` and cannot be removed from their own agenda.
6. Click the **X** beside a shared user to **Remove access**.
7. Click **Save Session** to persist content and sharing changes.

Sharing behavior:

- Sharing rows are stored in the `Sharing` tab.
- New shares can trigger **Agenda shares** email notifications.
- Shared agendas appear in the recipient's **Shared With Me** section.

### Profile, Settings, and Dark Mode

#### My Profile

Open the lower-left user menu and click **Profile**.

- The modal title is **My Profile**.
- It shows avatar/profile image, **Name**, and **Email**.
- Click **Edit Profile** to edit the form.
- Click **Save Changes** to persist the display name and refresh profile-photo URL if needed. The current backend does not update the login email even though the email input appears in the form.
- Click **Cancel** while editing, or **Close** when not editing.

#### Settings

Open the lower-left user menu and click **Settings**.

**Font Size** controls:

- **Decrease** button reduces font scale by 5%.
- Range slider supports 85% to 130% in 5% increments.
- **Increase** button increases font scale by 5%.
- Label shows the percent and a friendly label such as `Smaller`, `Default`, or `Larger`.

**Notifications** toggles:

- **Task assignments**
- **Task completion** (only shown to users who have direct reports; notifies managers when a report completes a task)
- **Comments and mentions**
- **Due-date reminders**
- **Weekly digest**
- **Agenda shares**

Footer buttons:

- **Cancel** closes without saving the modal state.
- **Save Settings** persists notification and font-size settings.

#### Dark Mode

- Open the lower-left user menu.
- Click **Dark Mode** to toggle between light and dark UI.
- The current menu status displays `On` or `Off`.
- Dark mode is a local UI state in the current browser session; user settings persistence is used for font size and notifications.

---

## Features

### Project Management

✅ **Create & Track Projects**
- Required title plus optional description and due date.
- Status options: `Not Started`, `In Progress`, `Completed`, `Delayed`.
- Project color schemes: `suu_red` (default), `sunset_orange`, `amber_gold`, `emerald_green`, `ocean_teal`, `sky_blue`, `deep_indigo`, `soft_violet`, `rose_pink`, `slate_gray`, `pearl_white` — 11 curated options.
- Auto-populated creation date and creator tracking.
- Project creator is automatically assigned to the project.
- Project columns can be reordered on the board.

✅ **Organize Work with Tasks**
- Create tasks inside projects with the **+ Add Task** button.
- Task status options: `Not Started`, `Upcoming`, `Review`, `In Progress`, `Ongoing`, `On Hold`, `Cancelled`, `Complete`.
- New tasks always begin as `Not Started`.
- Set optional priority: `High`, `Medium`, `Low`, or no priority.
- Due date management with visual indicators and due-tomorrow emphasis.
- Multiple assignees per task; at least one assignee is required by backend validation.
- Task cards can be dragged between projects.
- Auto-track completion timestamp and completing user when marked `Complete`.

✅ **Subtasks**
- Add subtasks while creating a task or later from **Task Details**.
- Toggle each subtask between `Incomplete` and `Complete`.
- Edit, delete, and reorder subtasks in task edit mode.

### Meeting Agendas

✅ **Recurring Agenda Sessions**
- Create agendas from **Meeting Agendas** with **New Agenda**.
- Each agenda is a recurring template; add a new **session** each time the team meets.
- The latest session date and section count are shown on agenda cards.
- Navigate past sessions with **← Older** / **Newer →** controls; past sessions are read-only.
- Agenda creators can create a new session (blank or copied from the previous session), delete sessions, and edit the current session's content.
- Organize session content with headers; add free-text items with **+ Text Item**; embed task references with **+ Link Task**.
- Open linked tasks using **View**.

✅ **Secure Sharing**
- Share agendas with specific SOAR users through **Share** → **Share Agenda**.
- Recipients see agendas in **Shared With Me**.
- Opt-in email notifications are sent when a new agenda is shared.

### Collaboration

✅ **Task Comments & Mentions**
- Add comments to task cards.
- Type `@` to select mention suggestions.
- Resolve comments to hide them from the active list.
- Delete comments permanently.
- View comment timestamps and authorship.

✅ **Smart Notifications**
1. **Task assignments**: users can be notified when assigned to a task.
2. **Comments and mentions**: users can be notified when mentioned in task comments.
3. **Task completion**: managers can be notified when assigned tasks are completed by their reports.
4. **Due-date reminders**: users can be alerted for open tasks due today or tomorrow.
5. **Weekly digest**: weekly summary of open assigned tasks.
6. **Agenda shares**: users can be notified when a teammate shares a meeting agenda.
7. **Account-created manager notice**: a manager can be notified when a report creates an account.

### User Management

✅ **Team Hierarchy**
- Onboard users with email, display name, and optional manager.
- Manager relationships use `Manager_ID`.
- Task assignment permissions allow self and reporting-tree users.
- Supervisor Tools visibility depends on direct reports.

✅ **User Profiles**
- Sync profile photos from Google account when possible.
- View email and edit display name in **My Profile**. The active backend identifies the user by signed-in Google email and does not persist profile email changes.
- Manage notification preferences per user.

### Accessibility & Personalization

✅ **Dark/Light Mode**
- Toggle from the lower-left user menu using **Dark Mode**.
- Displays current state as `On` or `Off`.

✅ **Font Scaling**
- Adjust font size from 85% to 130% in 5% increments.
- Use **Decrease**, slider, or **Increase** in **Settings**.

✅ **Notification Control**
- Toggle each notification type in **Settings**.
- Stored per user by email in Script Properties.

### Views & Navigation

✅ **Project Board**
- Visual project columns with task cards.
- Drag-and-drop project reorder and task move.
- Status dropdown on each task card.
- Color-coded project card styling.
- Assignee avatars on task cards.
- Comment button and unresolved-comment count.

✅ **Supervisor Tools**
- Select direct reports and view their open assigned tasks grouped by project.

✅ **Calendar**
- Month view for project due dates and task due dates.
- Click entries to open details.

✅ **Past Assignments**
- Completed assigned tasks.
- Permanent deletion for completed tasks.

✅ **Detail Modals**
- **Task Details**: full task editing, subtask management, assignee picker, completion, deletion.
- **Project Details**: project fields, creator display, color scheme, edit/delete actions.
- **Add Task**: task creation form.
- **New Project**: project creation form.
- **Comments**: task comment form and active comment list.
- **My Profile**: account profile viewer/editor. The UI shows **Name** and **Email**; the current backend persists the name and profile-photo refresh, but does not change the login email.
- **Settings**: font size and notification preferences.

✅ **SOAR Assistant**
- Optional Gemini-backed chat assistant in the bottom-right corner.
- Uses `Tutorial.html` plus current tab/direct-report context.

---

## Architecture

### System Overview

```
┌─────────────────────────────────────────────────────────┐
│                 Browser (Client Layer)                  │
│  Vue 3 SPA                                              │
│  • Project board, supervisor tools, calendar, agendas   │
│  • Modals, forms, settings, comments, assistant chat    │
│  • State management for users, projects, tasks, etc.    │
│  • Data sync via global version hashing + local cache   │
└────────────────────┬────────────────────────────────────┘
                     │ google.script.run calls
┌────────────────────┴────────────────────────────────────┐
│         Google Apps Script V8 Runtime                   │
│                                                         │
│  Web App Layer                                          │
│  • Code.js → doGet(), include()                         │
│  • Bootstrap.js → getInitialPayload()                   │
│                                                         │
│  Business Logic Services                                │
│  • Users.js → User CRUD + profile updates               │
│  • Projects.js → Project lifecycle + ordering           │
│  • Tasks.js → Task operations + status/order changes    │
│  • Subtasks.js → Subtask CRUD + ordering                │
│  • Comments.js → Task comments + mention extraction     │
│  • Agendas.js → Agenda + session CRUD + sharing logic   │
│  • Notifications.js → email notification types          │
│  • Settings.js → User preference persistence            │
│  • Chat.js → Optional Gemini assistant bridge           │
│                                                         │
│  Data Access Layer                                      │
│  • DataStore.js → sheet reads/writes, IDs, cache, hash   │
│                                                         │
│  Utilities Layer                                        │
│  • Utilities.js → email, dates, profile photos, mail    │
└────────────────────┬────────────────────────────────────┘
                     │ SpreadsheetApp / DriveApp / MailApp
┌────────────────────┴────────────────────────────────────┐
│        Google Sheets (Data Persistence Layer)           │
│  Tabs: Users | Projects | Tasks | Subtasks | Comments   │
│        Assignments | Agendas | Sharing | Sessions│
└─────────────────────────────────────────────────────────┘
```

### Frontend Architecture (Vue 3)

- **Reactive Data**: Refs for `users`, `projects`, `tasks`, `subtasks`, `assignments`, `comments`, `agendas`, `agendaShares`, and `agendaSessions`.
- **State Management**: Computed properties for current user, visibility, assignee summaries, project summaries, calendar entries, agenda ownership, shared agendas, and supervisor selections.
- **Data Sync**: `getGlobalVersionHash()` checks whether cached payload data is still current; if not, `getInitialPayload()` reloads app data.
- **UI Framework**: Tailwind CSS, Font Awesome icons, SortableJS/Vue Draggable, and Marked for assistant markdown rendering.

### Backend Communication

- **Web App Access**: Manifest uses `ANYONE` access and `USER_ACCESSING` execution.
- **Transport**: Client code calls server functions with `google.script.run`.
- **Response Shape**: Most mutation functions return JSON strings with `success` and optional `error`.
- **Locking**: `LockService` prevents race conditions during ID generation.

### Caching Strategy

| Level | Scope | TTL | Purpose |
|-------|-------|-----|---------|
| `REQUEST_CACHE` | Single Apps Script execution | Execution lifetime | Avoid repeated sheet reads in one request |
| `CacheService` | Script cache | 300 seconds | Cache `Users` table |
| `ScriptProperties` | Persistent | No TTL | ID counters, user settings, app data version |
| Browser storage | Current browser | Session/local storage | Logo processing and initial payload cache |

**Invalidation**: Every table write should call `invalidateTableCache(tableName)`, which clears request cache, removes the user cache when needed, and bumps `soar_data_version:last_updated`.

### Permission Model

- **Identity**: Current user is determined from the signed-in Google account email.
- **Manager Hierarchy**: Users have optional `Manager_ID`.
- **Task Assignment Permissions**: A user can assign tasks to themselves and users in their reporting tree. Supervisor selection UI lists direct reports.
- **Project Visibility**: Users see assigned projects and projects containing open tasks assigned to them.
- **Creator Ownership**: Project creators are automatically assigned to their projects.

### Data Flow Example: Creating a Task

```
1. User clicks "+ Add Task" at the bottom of a project column.
2. Vue opens the "Add Task" modal for that project.
3. User fills Task Title, Due Date, Priority, Assigned To, Description, and optional Subtasks.
4. User clicks "Create Task".
5. Vue calls createTask(projectId, taskInput) with assignee User_IDs and subtask titles.
6. Apps Script validates project, title, assignees, assignee permissions, priority, and date.
7. Apps Script generates Task_ID (T-00000001 style) with LockService.
8. Apps Script appends a row to Tasks.
9. Apps Script appends one Assignments row per assignee.
10. Apps Script optionally appends Subtasks rows with S- IDs and Incomplete status.
11. Apps Script invalidates table caches and bumps the global data version.
12. Apps Script sends task-assignment notifications according to recipient settings.
13. Apps Script returns the created task, assignments, and subtasks.
14. Vue updates local state, closes the modal, and re-renders the board.
```

---

## Data Model

### Entity Relationships

```
User (U-00000001)
├── manages → User[] via Users.Manager_ID
├── creates → Project[] via Projects.Creator_ID
├── creates → Agenda[] via Agendas.Creator_ID
├── completes → Task[] via Tasks.Completed_By
└── comments on → Task comments via Comments.Commenter_ID

Project (P-00000001)
├── contains → Task[] via Tasks.Project_ID
├── assigned to → User[] via Assignments.Assignment_ID = Project_ID
└── has visual color via Projects.Color_Scheme

Task (T-00000001)
├── belongs to → Project via Tasks.Project_ID
├── assigned to → User[] via Assignments.Assignment_ID = Task_ID
├── contains → Subtask[] via Subtasks.Task_ID
├── receives → Comment[] via Comments.Topic_ID
└── completed by → User via Tasks.Completed_By

Subtask (S-00000001)
├── belongs to → Task via Subtasks.Task_ID
└── has Status = Incomplete or Complete

Comment (C-00000001)
├── on → Task via Comments.Topic_ID
├── by → User via Comments.Commenter_ID
└── mentions → User[] extracted from content handles

Agenda (A-00000001)
├── created by → User via Agendas.Creator_ID
├── contains → AgendaSession[] via Sessions.Agenda_ID
└── shared with → User[] via Sharing

AgendaSession (AS-00000001)
├── belongs to → Agenda via Sessions.Agenda_ID
└── stores session content as Content_JSON

Assignment
├── Assignment_ID = Task_ID or Project_ID
└── Assignee_ID = User_ID
```

### ID Generation Pattern

Primary keys are human-readable, auto-incrementing, and zero-padded:

- **Users**: `U-00000001`, `U-00000002`, ...
- **Projects**: `P-00000001`, `P-00000002`, ...
- **Tasks**: `T-00000001`, `T-00000002`, ...
- **Subtasks**: `S-00000001`, `S-00000002`, ...
- **Comments**: `C-00000001`, `C-00000002`, ...
- **Agendas**: `A-00000001`, `A-00000002`, ...
- **Agenda Sessions**: `AS-00000001`, `AS-00000002`, ...

Generated by Apps Script with synchronized locking to prevent race conditions.

### Data Dictionary

> **Notes**
> - **Can Be Null = No** means the field is required at creation time unless auto-filled by the system.
> - Foreign-key relationships are enforced by application logic, not by Google Sheets.
> - Dates are stored as Sheets date/datetime values and serialized for the client.

#### Users

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `User_ID` | String, auto-increment | Format: `U-00000000` | No |
| `Email` | Email address | Unique login identifier | No |
| `Name` | String | Display name | No |
| `Manager_ID` | String, User_ID reference | Manager's `User_ID` | Yes |
| `Profile_Pic_Url` | URL | Google Account profile photo | Yes |

#### Projects

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Project_ID` | String, auto-increment | Format: `P-00000000` | No |
| `Project_Title` | String | Display title | No |
| `Description` | String | Goals and scope | Yes |
| `Status` | String | `Not Started`, `In Progress`, `Completed`, `Delayed` | No |
| `Created_Date` | DateTime | Auto-populated at creation | No |
| `Due_Date` | Date | Planned completion date | Yes |
| `Creator_ID` | String, User_ID reference | Project creator | No |
| `Color_Scheme` | String | One of 11 keys: `suu_red` (default), `sunset_orange`, `amber_gold`, `emerald_green`, `ocean_teal`, `sky_blue`, `deep_indigo`, `soft_violet`, `rose_pink`, `slate_gray`, `pearl_white` | Yes |

#### Tasks

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Task_ID` | String, auto-increment | Format: `T-00000000` | No |
| `Project_ID` | String, Project_ID reference | Parent project | No |
| `Task_Title` | String | Display title | No |
| `Description` | String | Task details | Yes |
| `Status` | String | `Not Started`, `Upcoming`, `Review`, `In Progress`, `Ongoing`, `On Hold`, `Cancelled`, `Complete` | No |
| `Priority` | String | `High`, `Medium`, `Low`, or blank | Yes |
| `Created_Date` | DateTime | Auto-populated at creation | No |
| `Due_Date` | Date | Planned completion date | Yes |
| `Completed_By` | String, User_ID reference | User who completed the task | Yes |
| `Completed_At` | DateTime | Completion timestamp | Yes |

#### Subtasks

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Subtask_ID` | String, auto-increment | Format: `S-00000000` | No |
| `Task_ID` | String, Task_ID reference | Parent task | No |
| `Subtask_Title` | String | Display title | No |
| `Status` | String | `Incomplete` or `Complete` | No |

#### Comments

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Comment_ID` | String, auto-increment | Format: `C-00000000` | No |
| `Topic_ID` | String | Active UI uses `Task_ID` | No |
| `Topic_Type` | String | Active UI writes `Task` | No |
| `Commenter_ID` | String, User_ID reference | Comment author | No |
| `Content` | String | Comment text with optional `@handle` mentions | No |
| `Timestamp` | DateTime | Auto-populated at creation | No |
| `Is_Resolved` | Boolean | Whether comment is hidden from active list | No |

#### Assignments

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Assignment_ID` | String | Task_ID or Project_ID | No |
| `Assignee_ID` | String | User_ID of assigned user | No |

#### Agendas

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Agenda_ID` | String, auto-increment | Format: `A-00000000` | No |
| `Title` | String | Agenda title | No |
| `Creator_ID` | String, User_ID reference | Agenda creator | No |
| `Created_Date` | DateTime | Auto-populated at creation | No |
| `Description` | String | Optional description shown on the agenda card and editable in the editor | Yes |
| `Content_JSON` | JSON string | Legacy field — retained for migration purposes; session content is now stored in `Sessions` | Yes |
| `Agenda_Date` | Date | Legacy field — retained for migration purposes; session dates are now stored in `Sessions` | Yes |

#### Sessions

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Session_ID` | String, auto-increment | Format: `AS-00000000` | No |
| `Agenda_ID` | String, Agenda_ID reference | Parent agenda | No |
| `Session_Date` | Date | Date of this meeting session | Yes |
| `Content_JSON` | JSON string | Array of headers containing text items and linked task items | Yes |
| `Created_Date` | DateTime | Auto-populated at creation | No |

#### Sharing

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Agenda_ID` | String, Agenda_ID reference | Shared agenda | No |
| `User_ID` | String, User_ID reference | User with access | No |

---

## Configuration

### appsscript.json

The manifest file defines permissions, runtime, and deployment settings:

```json
{
  "timeZone": "America/Denver",
  "dependencies": {},
  "oauthScopes": [
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/drive.metadata.readonly"
  ],
  "webapp": {
    "access": "ANYONE",
    "executeAs": "USER_ACCESSING"
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

**Key settings**:

- **timeZone**: Used for due date handling and digest/reminder scheduling.
- **script.external_request**: Needed by the optional Gemini assistant.
- **spreadsheets**: Needed for all Google Sheets storage.
- **userinfo.email/profile**: Needed to identify users and fetch profile info.
- **script.send_mail**: Needed for notification emails.
- **drive.metadata.readonly**: Used to read spreadsheet last-updated metadata for version hashing.
- **webapp.access**: `ANYONE` allows the deployed URL to load; users still identify through Google account email.
- **webapp.executeAs**: `USER_ACCESSING` runs as the accessing user.
- **runtimeVersion**: `V8` required.

### Script Properties

| Property | Required | Purpose |
|---|---:|---|
| `GEMINI_API_KEY` | Optional | Enables SOAR Assistant calls to Gemini. |
| `soar_data_version:last_updated` | Auto-created | App data version for cache invalidation. |
| `soar_next_id:{Sheet}:{Prefix}` | Auto-created | ID counters for generated IDs. |
| `soar_user_settings:{email}` | Auto-created | User font size and notification settings. |

### Spreadsheet Setup

Each sheet tab requires specific column headers. Keep these names exact.

**Users**:
```
User_ID | Email | Name | Manager_ID | Profile_Pic_Url
```

**Projects**:
```
Project_ID | Project_Title | Description | Status | Created_Date | Due_Date | Creator_ID | Color_Scheme
```

**Tasks**:
```
Task_ID | Project_ID | Task_Title | Description | Status | Priority | Created_Date | Due_Date | Completed_By | Completed_At
```

**Subtasks**:
```
Subtask_ID | Task_ID | Subtask_Title | Status
```

**Comments**:
```
Comment_ID | Topic_ID | Topic_Type | Commenter_ID | Content | Timestamp | Is_Resolved
```

**Assignments**:
```
Assignment_ID | Assignee_ID
```

**Agendas**:
```
Agenda_ID | Title | Creator_ID | Created_Date | Content_JSON | Description | Agenda_Date
```

**Sessions**:
```
Session_ID | Agenda_ID | Session_Date | Content_JSON | Created_Date
```

**Sharing**:
```
Agenda_ID | User_ID
```

---

## API Reference

Most server functions return a JSON string. The client parses the response with `parseRunResponse()`.

### User Functions

#### `addUser(userInput)`
Creates a user record. The onboarding **Create Account** modal calls this function with the signed-in email.

**Parameters**:
- `userInput.email` (string): required user email.
- `userInput.name` (string): required display name.
- `userInput.managerId` (string, optional): selected manager `User_ID`.

**Returns**: `{success: true, user: {...}, created: true}` for a new user, `{success: true, user: {...}, created: false}` for an existing user, or `{success: false, error: "..."}`

#### `updateCurrentUserProfile(profileInput)`
Updates the logged-in user's profile.

**Parameters**:
- `profileInput.name` (string): display name.
- `profileInput.email` (string): email value from profile form.

**Returns**: `{success: true, user: {...}}`

### Project Functions

#### `createProject(projectInput)`
Creates a new project and assigns it to the creator.

**Parameters**:
- `projectInput.projectTitle` (string): required project title.
- `projectInput.description` (string, optional): project description.
- `projectInput.status` (string): `Not Started`, `In Progress`, `Completed`, or `Delayed`.
- `projectInput.dueDate` (string, optional): `YYYY-MM-DD`.
- `projectInput.colorScheme` (string, optional): one of the 11 color scheme keys — `suu_red`, `sunset_orange`, `amber_gold`, `emerald_green`, `ocean_teal`, `sky_blue`, `deep_indigo`, `soft_violet`, `rose_pink`, `slate_gray`, `pearl_white`; defaults to `suu_red`.

**Returns**: `{success: true, project: {...}, assignment: {...}}`

#### `updateProject(projectId, projectInput)`
Updates project metadata.

**Parameters**:
- `projectId` (string): Project_ID to update.
- `projectInput.projectTitle`, `description`, `status`, `dueDate`, `colorScheme`.

**Returns**: `{success: true, project: {...}}`

#### `deleteProject(projectId)`
Deletes the project row, tasks in that project, and assignment rows for those deleted tasks. It does not currently remove the project creator assignment row from `Assignments`.

**Returns**: `{success: true, projectId: "P-00000001"}`

#### `reorderProjects(orderedProjectIds)`
Persists project display order by rewriting project rows in the supplied order. Deprecated in favor of per-user ordering via `saveUserSortOrder()`.

### Task Functions

#### `createTask(projectId, taskInput)`
Creates a new task inside a project.

**Parameters**:
- `projectId` (string): parent Project_ID.
- `taskInput.taskTitle` (string): required task title.
- `taskInput.description` (string, optional): task description.
- `taskInput.priority` (string, optional): `High`, `Medium`, or `Low`.
- `taskInput.dueDate` (string, optional): `YYYY-MM-DD`.
- `taskInput.assigneeIds` (array): required list of `User_ID` values.
- `taskInput.subtasks` (array, optional): list of subtask title strings.

**Returns**: `{success: true, task: {...}, assignments: [...], subtasks: [...]}`

#### `updateTask(taskId, taskInput)`
Updates task metadata, project, assignees, and newly added subtasks.

**Parameters**:
- `taskId` (string): Task_ID to update.
- `taskInput.taskTitle`, `description`, `status`, `priority`, `dueDate`, `projectId`, `assigneeIds`, `newSubtasks`.

**Returns**: `{success: true, task: {...}, assignments: [...], newSubtasks: [...]}`

#### `updateTaskStatus(taskId, newStatus)`
Updates only a task's status. Completion fields are populated when the new status is `Complete`.

#### `completeTask(taskId)`
Marks a task as completed and records completion timestamp and completing user.

**Net Effect**: Sets `Status = "Complete"`, `Completed_By = current_user`, `Completed_At = now()`.

#### `deleteTask(taskId)`
Deletes a task and related assignment rows.

#### `deleteTask(taskId)` from **Past Assignments**
The **Delete Permanently** button in **Past Assignments** calls the same backend `deleteTask(taskId)` function after a stronger confirmation message.

#### `moveTaskToProject(taskId, newProjectId)`
Moves a task to another project by updating its `Project_ID`. Enforces that the calling user is the creator of both the source and target project. Returns an error if the permission check fails.

#### `saveUserSortOrder(entityType, orderedIds)`
Persists a user's personal display order for `'projects'` or `'tasks'` to `PropertiesService`. Only affects the calling user's view — other users' orders are unchanged.

#### `updateTaskProjectAndOrder(taskId, newProjectId, orderedTaskIdsInProject)`
Legacy function that moves a task to another project and reorders the destination project's sheet rows. Superseded by `moveTaskToProject()` + `saveUserSortOrder()`.

#### `purgeCompletedTasksPastDue()`
Deletes completed tasks whose due dates have passed, plus their assignment rows.

### Subtask Functions

#### `addSubtask(taskId, subtaskTitle)`
Adds an `Incomplete` subtask to an existing task.

#### `updateSubtaskStatus(subtaskId, isComplete)`
Sets a subtask status to `Complete` or `Incomplete`.

#### `updateSubtaskTitle(subtaskId, title)`
Updates an existing subtask title.

#### `deleteSubtask(subtaskId)`
Deletes a subtask from a task.

#### `reorderSubtasks(taskId, orderedSubtaskIds)`
Persists subtask display order for a task.

### Comment Functions

#### `getCommentsByTopic(topicId)`
Returns unresolved comments for a task topic.

#### `addComment(topicId, commentInput)`
Creates a task comment and sends mention notifications.

**Parameters**:
- `topicId` (string): active UI uses Task_ID.
- `commentInput.content` (string): comment text with optional `@handle` mentions.

**Returns**: `{success: true, comment: {...}}`

#### `deleteComment(commentId)`
Deletes a comment.

#### `resolveComment(commentId)`
Marks a comment as resolved.

### Agenda Functions

#### `createAgenda(title)`
Creates a new blank agenda template (no sessions).

**Returns**: `{success: true, agenda: {...}}`

#### `updateAgenda(agendaId, title, description, sharedUserIds)`
Updates an agenda's title, description, and sharing permissions.

**Parameters**:
- `agendaId` (string): Agenda_ID.
- `title` (string): agenda title.
- `description` (string): agenda description.
- `sharedUserIds` (array): `User_ID` values with access.

**Returns**: `{success: true, agenda: {...}, shares: [...]}`

#### `deleteAgenda(agendaId)`
Deletes an agenda, all its sessions, and associated sharing permissions.

#### `createAgendaSession(agendaId, sessionDate, contentJson)`
Creates a new session for the agenda. Only the agenda creator may call this.

**Parameters**:
- `agendaId` (string): Agenda_ID.
- `sessionDate` (string): `YYYY-MM-DD` meeting date.
- `contentJson` (string): stringified JSON array of agenda sections/items.

**Returns**: `{success: true, session: {...}}`

#### `updateAgendaSession(sessionId, sessionDate, contentJson)`
Updates an existing session's date and content.

**Parameters**:
- `sessionId` (string): Session_ID.
- `sessionDate` (string): `YYYY-MM-DD` meeting date.
- `contentJson` (string): stringified JSON array of agenda sections/items.

**Returns**: `{success: true, session: {...}}`

#### `getAgendaSessions(agendaId)`
Returns all sessions for an agenda sorted by date descending.

**Returns**: `{success: true, sessions: [...]}`

#### `deleteAgendaSession(sessionId)`
Deletes a single agenda session.

**Returns**: `{success: true, sessionId: "..."}`

### Settings Functions

#### `persistCurrentUserSettings(settingsInput)`
Persists current user's font scale and notification settings.

#### `getUserSettingsByEmail(email)`
Loads stored settings or defaults.

### Notification Functions

These functions are normally called internally:

- `sendTaskAssignmentNotifications(task, assigneeIds, assignedByUserId)`
- `sendMentionNotifications(comment, topicId, commenter, mentionedUsers)`
- `sendManagerTaskCompletedNotifications(task, assigneeIds, completedByUserId)`
- `sendManagerAccountCreatedNotification(createdUser)`
- `sendDueDateReminderNotifications()`
- `sendWeeklyDigestNotifications()`
- `sendAgendaShareNotifications(agenda, sharedUserIds, sharedByUserId)`

### Bootstrap, Versioning, and Chat

#### `getInitialPayload()`
Fetches initial app state.

**Returns**:
```json
{
  "currentUserEmail": "user@example.com",
  "currentUserExists": true,
  "requiresAccountSetup": false,
  "users": [],
  "projects": [],
  "tasks": [],
  "subtasks": [],
  "assignments": [],
  "agendas": [],
  "agendaShares": [],
  "agendaSessions": [],
  "comments": [],
  "currentUserSettings": {},
  "versionHash": "...",
  "lastUpdated": "..."
}
```

#### `getGlobalVersionHash()`
Returns a version hash based on spreadsheet metadata and app data version.

#### `askGeminiAssistant(conversationHistory, userContext)`
Calls Gemini with SOAR assistant instructions, `Tutorial.html`, recent conversation history, and current UI context. Requires Script Property `GEMINI_API_KEY`.

---

## Development

### Adding a New Feature

#### 1. Extend the Data Model

If your feature requires new data:
1. Create a new sheet tab for the entity.
2. Define column headers and update [Spreadsheet Setup](#spreadsheet-setup).
3. Add CRUD functions to a new service file, or extend the relevant existing service file.

#### 2. Add Backend Logic

Use the existing service pattern:

```javascript
function createNewThing(input) {
  try {
    const title = input && input.title ? input.title.toString().trim() : '';
    if (!title) throw new Error('Title is required.');

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NewThings');
    if (!sheet) throw new Error('NewThings sheet was not found.');

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndex = getHeaderIndex(headers);
    const row = new Array(headers.length).fill('');

    if (headerIndex.NewThing_ID !== undefined) row[headerIndex.NewThing_ID] = generateNextId('NewThings', 'N');
    if (headerIndex.Title !== undefined) row[headerIndex.Title] = title;
    if (headerIndex.Created_Date !== undefined) row[headerIndex.Created_Date] = new Date();

    appendRows(sheet, [row]);
    invalidateTableCache('NewThings');

    return JSON.stringify({ success: true });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message || 'Failed to create item.' });
  }
}
```

#### 3. Extend Frontend UI

In [App.js.html](App.js.html), add Vue state and methods as needed:

- new `ref()` data collections;
- computed properties for derived state;
- `google.script.run` calls;
- new view sections or modals in [Index.html](Index.html);
- navigation updates if the feature needs a new tab.

#### 4. Add Email Notifications (if applicable)

In [Notifications.js](Notifications.js), add notification helpers and check user preferences with `isNotificationEnabledForUser()` or `isNotificationEnabledForEmail()`.

#### 5. Test & Document

- Test with an Apps Script test deployment.
- Verify frontend calls and sheet writes.
- Check browser console and Apps Script **Executions** logs.
- Update this README and, if assistant behavior changes, [Tutorial.html](Tutorial.html).

### Code Organization

- **Utilities.js**: Shared helpers for identity, profile photos, email validation, date parsing, mail sending, and client serialization.
- **DataStore.js**: Data access layer, cache invalidation, version hashing, ID generation, row writes/deletes.
- **Users.js**: User creation, onboarding, profile updates, and user lookup.
- **Projects.js**: Project CRUD, color schemes, creator assignment, project ordering.
- **Tasks.js**: Task CRUD, status flow, assignment validation, task moving/order, completed-task purge.
- **Subtasks.js**: Subtask CRUD and ordering.
- **Comments.js**: Task comments, resolved state, mention extraction.
- **Agendas.js**: Agenda CRUD and sharing rows.
- **Notifications.js**: Email notification workflows.
- **Settings.js**: Per-user notification and font-scale settings.
- **Bootstrap.js**: Initial app payload.
- **Chat.js**: Optional Gemini assistant bridge.
- **Code.js**: Web app entry point (`doGet()`) and HTML include helper.
- **Index.html**: HTML template and UI markup.
- **App.js.html**: Vue app logic.
- **Tutorial.html**: Assistant tutorial content loaded by `Chat.js`.

### Running Tests

Currently, Apps Script testing is manual:

1. Deploy as test deployment: **Deploy** → **Test deployments**.
2. Open the URL in a browser.
3. Verify features end-to-end.
4. Check browser console for errors.
5. Review Apps Script logs in the **Executions** tab.

For repository-level validation, run a syntax-oriented check such as:

```bash
node --check /tmp/combined-soar-js.js
```

where `/tmp/combined-soar-js.js` is a temporary file built from the `.js` files after stripping Apps Script HTML wrappers if needed.

### Code Style & Conventions

- **Functions**: camelCase (for example, `createTask()`, `getInitialPayload()`).
- **Variables**: camelCase for locals; UPPER_CASE for constants.
- **Responses**: Return JSON strings with `{success: false, error: "message"}` on failures.
- **Validation**: Validate all inputs at function entry points.
- **Caching**: Call `invalidateTableCache(tableName)` after writes.
- **Imports**: Do not wrap imports in try/catch blocks.

---

## Troubleshooting

### App won't load / blank screen

**Causes & Solutions**:
1. **Deployment URL is wrong**: Verify you deployed as **Web app** and are using the current deployment URL.
2. **Apps Script hasn't mounted Vue**: Check browser console for JavaScript errors.
3. **Sheet tabs missing**: Verify all 8 tabs exist: `Users`, `Projects`, `Tasks`, `Subtasks`, `Comments`, `Assignments`, `Agendas`, `Sharing`.
4. **Sheet headers are wrong**: Verify exact headers in [Spreadsheet Setup](#spreadsheet-setup).
5. **Apps Script permissions not granted**: Refresh page and authorize requested scopes.

**Debug steps**:
- Open browser DevTools (F12).
- Check **Console** for JavaScript errors.
- Check **Network** for failed calls.
- Open Apps Script editor and check **Executions** logs.

### Create Account modal keeps appearing

**Possible Causes**:
1. The signed-in email is not in `Users.Email`.
2. Email has extra spaces or different casing in the sheet.
3. The web app is executing under a context that cannot read the current user's email.

**Solutions**:
- Confirm `Users` has an exact email for the signed-in account.
- Use the **Create Account** modal and click **Create Account**.
- Confirm manifest scopes include `userinfo.email`.

### Tasks or projects not appearing

**Possible Causes**:
1. You are seeing only assigned work by default.
2. The project is not assigned to you and contains no open tasks assigned to you.
3. The task is complete and moved to **Past Assignments**.
4. Sheet columns are wrong or missing.
5. Cache is stale.

**Solutions**:
- Check `Assignments` rows for your `User_ID`.
- Check whether the task status is `Complete`.
- Refresh the browser.
- Check Apps Script logs for errors.

### Cannot create a task

**Possible Causes**:
1. **Task Title** is blank.
2. No assignee was selected under **Assigned To**.
3. You selected a user outside your permitted reporting tree.
4. The target project was deleted or the `Project_ID` is invalid.

**Solutions**:
- Add a title.
- Select yourself or an allowed report as assignee.
- Ask a manager/admin to adjust `Manager_ID` values if permissions are wrong.

### Email notifications not sending

**Possible Causes**:
1. `script.send_mail` scope was not authorized.
2. User disabled the relevant notification in **Settings**.
3. Recipient email is invalid.
4. Apps Script daily email quotas were reached.

**Debug**:
- Check Apps Script **Executions** logs.
- Confirm notification settings.
- Confirm `Users.Email` values are valid.

### SOAR Assistant says it is not configured

**Cause**: Script Property `GEMINI_API_KEY` is missing.

**Solution**:
1. Open Apps Script **Project Settings**.
2. Add Script Property `GEMINI_API_KEY`.
3. Save and retry the chat.

### "You do not have permission to access this file"

**Cause**: Wrong deployment URL, deleted Apps Script project, or missing access to the spreadsheet/script.

**Solution**:
1. Verify you have edit access to the Google Sheet.
2. Open **Extensions** → **Apps Script** from the correct sheet.
3. Deploy a fresh web app.
4. Use the new URL.

### Slow performance / high latency

**Causes**:
1. Large Google Sheets data volume.
2. Many concurrent users.
3. Rapid writes causing frequent cache invalidation.
4. Apps Script quotas or cold starts.

**Solutions**:
1. Archive old completed projects/tasks if sheets grow large.
2. Batch operations when possible.
3. Avoid unnecessary full reloads.
4. Review [Caching Strategy](#caching-strategy).

### Mention extraction not working

**Problem**: A typed mention is not recognized or notified.

**Causes**:
1. Mentions must use handles like `@jane` or `@first.last`; display names with spaces are not the backend mention format.
2. The user is not in the `Users` sheet.
3. The mentioned user disabled **Comments and mentions**.

**Solution**:
- Use the `@` suggestion dropdown when writing the comment.
- Ensure the user exists in SOAR.
- Check notification settings.

### Permissions error: "Cannot assign task to this user"

**Cause**: You selected an assignee outside your reporting tree.

**Why**: Task assignment permissions allow:
- yourself;
- your direct reports;
- indirect reports below your direct reports.

**Solution**:
- Assign the task to yourself or someone in your reporting tree.
- Ask an admin to update `Manager_ID` values.
- Have the appropriate manager create or edit the task.

### Generated IDs reset or duplicated IDs

**Cause**: Script Properties counter corruption or manual sheet edits.

**Workaround**:
1. Open Apps Script editor.
2. Open a temporary function or console context.
3. Delete affected ID counter properties, for example:
   ```javascript
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Users:U');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Projects:P');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Tasks:T');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Subtasks:S');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Comments:C');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Agendas:A');
   ```
4. The next mutation re-seeds from existing sheet IDs.

---

## Support & Contributing

### Reporting Issues

Found a bug? Have a feature request?
1. Check [Troubleshooting](#troubleshooting) first.
2. Open an issue on [GitHub](https://github.com/nathanwiggins/soar/issues).
3. Include steps to reproduce, expected vs actual behavior, browser/OS, and relevant logs.

### Contributing

To add a feature or fix a bug:
1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-feature`.
3. Make changes following [Code Style & Conventions](#code-style-conventions).
4. Test end-to-end in a deployed Apps Script app.
5. Update this README if behavior changes.
6. Update [Tutorial.html](Tutorial.html) if assistant-facing guidance changes.
7. Submit a pull request with a description of changes.

### Roadmap

Potential future enhancements:
- [ ] File attachments on tasks/comments
- [ ] Recurring tasks
- [ ] Time tracking / Kanban burn-down charts
- [ ] Mobile app
- [ ] Slack integration for notifications
- [ ] Saved filters
- [ ] Bulk import from CSV

---

## License

Soar is provided as-is for educational and professional use within Google Workspace environments.

---

## Credits

Built by RAT using Google Apps Script, Vue.js, and Tailwind CSS.
