# Soar — Project Management for Google Workspace

**Soar** is a lightweight, Google-native project management system built on **Google Apps Script** and **Google Sheets**. It provides team collaboration features including project tracking, task management, comments with mentions, and intelligent email notifications—all without external infrastructure.

**Perfect for**: Small to mid-sized teams already using Google Workspace who want project management without complex setup or external dependencies.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [Features](#features)
3. [Architecture](#architecture)
4. [Data Model](#data-model)
5. [Configuration](#configuration)
6. [API Reference](#api-reference)
7. [Development](#development)
8. [Troubleshooting](#troubleshooting)

---

## Quick Start

### Prerequisites

- Google account with Apps Script access (part of Google Workspace)
- Permission to create a new Google Sheet
- Permission to deploy Apps Script web apps

### Deployment (5 minutes)

1. **Create a new Google Sheet**
   - Go to [sheets.google.com](https://sheets.google.com)
   - Click **New** → **Blank spreadsheet**
   - Name it "Soar" (or your preferred name)
   - Note the Sheet ID from the URL: `https://docs.google.com/spreadsheets/d/{SHEET_ID}/...`

2. **Create the data structure**
   - Create 5 new sheet tabs with these exact names (right-click sheet tab → Insert sheet):
     - `Users`
     - `Projects`
     - `Tasks`
     - `Comments`
     - `Assignments`
   - Add header rows to each tab (see [Data Model](#data-model) section for column names)

3. **Create Apps Script project**
   - In your Google Sheet, go to **Extensions** → **Apps Script**
   - A new Apps Script project will open
   - Delete any default content in `Code.gs`

4. **Add the source code**
   - Copy all files from this repository (in order):
     - [Utilities.js](Utilities.js)
     - [DataStore.js](DataStore.js)
     - [Users.js](Users.js)
     - [Projects.js](Projects.js)
     - [Tasks.js](Tasks.js)
     - [Comments.js](Comments.js)
     - [Notifications.js](Notifications.js)
     - [Settings.js](Settings.js)
     - [Bootstrap.js](Bootstrap.js)
     - [Code.js](Code.js) (paste into existing `Code.gs`)
     - [Index.html](Index.html) → Create as HTML file in Apps Script editor
     - [App.js.html](App.js.html) → Create as HTML file in Apps Script editor
   - Update `appsscript.json` with settings (see [Configuration](#configuration))

5. **Deploy as web app**
   - Click **Deploy** → **New deployment**
   - Type: Select **Web app**
   - Execute as: Your Google account
   - Who has access: **Anyone**
   - Click **Deploy**
   - You'll get a URL like `https://script.google.com/macros/d/{DEPLOYMENT_ID}/userweb`

6. **Initialize the first user**
   - Open the deployment URL in a browser
   - System will prompt you to create a user account
   - Enter your email, name, and (optionally) manager
   - ✅ You're ready to go!

### First Steps in the App

- **Create a project**: Click "New Project" button, fill details, choose status
- **Create tasks**: Within a project, add tasks with priority and assignees
- **Assign tasks**: Select team members from your manager hierarchy
- **Collaborate**: Add comments to tasks/projects using `@username` mentions
- **Customize**: Adjust font size and notification preferences in Settings

---

## Features

### Project Management

✅ **Create & Track Projects**
- Title, description, status, and due date
- Status workflow: `Not Started` → `In Progress` → `Completed` or `Delayed`
- Auto-populated creation date and creator tracking
- Assign team members to projects

✅ **Organize Work with Tasks**
- Create tasks within projects
- Set priority: `High`, `Medium`, `Low`
- Track status independently from parent project
- Due date management with visual indicators
- Auto-track completion timestamp and completing user
- Multiple assignees per task

### Collaboration

✅ **Comments & Mentions**
- Add comments to any task or project
- `@mention` users to notify them directly
- Resolve comments to keep threads organized
- View comment history and authorship

✅ **Smart Notifications** (5 types)
1. **Task Assignment**: Notified when assigned to a task
2. **Mentions**: Tagged when mentioned in comments
3. **Task Completion**: Managers notified when assigned tasks are completed
4. **Due Date Reminders**: Alerts for tasks due today or tomorrow
5. **Weekly Digest**: Sunday email summary of open assigned tasks

### User Management

✅ **Team Hierarchy**
- Onboard users with email and display name
- Manager relationships for permission enforcement
- Role-based task assignment (assign only to subordinates or self)

✅ **User Profiles**
- Sync profile photos from Google account
- Edit display name
- Manage notification preferences per user

### Accessibility & Personalization

✅ **Dark/Light Mode**
- Toggle between dark and light themes
- Preference saved per user

✅ **Font Scaling**
- Adjust font size from 85% to 130% in 5% increments
- Improves readability for all users

✅ **Notification Control**
- Toggle each notification type (assignments, mentions, due dates, weekly digest)
- User-controlled preference per notification type

### Views & Navigation

✅ **Kanban Board**
- Visual task management by status columns
- Drag-and-drop to change task status
- Color-coded priority indicators
- Assignee avatars on task cards
- Quick action buttons (complete, comment)

✅ **All Work Items List**
- Table view of tasks and projects
- Sortable columns
- Inline edit and delete actions
- Filter by project

✅ **Detail Modals**
- **Task Details**: Full task editing, assignee picker, embedded comments
- **Project Details**: Project info, member assignments, task list
- **Create Modals**: Quick forms for new tasks and projects
- **User Profile**: Edit name, view email, see profile photo
- **Settings**: Font size, theme, notification preferences

---

## Architecture

### System Overview

```
┌─────────────────────────────────────────────────────────┐
│                 Browser (Client Layer)                  │
│  Vue 3 SPA (3000+ lines of interactive UI code)         │
│  • Kanban board, modals, forms, settings                │
│  • Real-time data sync via version hashing              │
│  • State management (users, projects, tasks, comments)  │
└────────────────────┬────────────────────────────────────┘
                     │ AJAX/Fetch (JSON over HTTPS)
┌────────────────────┴────────────────────────────────────┐
│         Google Apps Script V8 Runtime                   │
│        (Backend Logic Layer - 7 modules)                │
│                                                          │
│  ┌──────────────────────────────────────────────────┐   │
│  │ Web App Layer (Code.js)                          │   │
│  │ • doGet() serves HTML shell                     │   │
│  │ • getInitialPayload() loads app state           │   │
│  └──────────────────────────────────────────────────┘   │
│                                                          │
│  ┌──────────────────────────────────────────────────┐   │
│  │ Business Logic Services                          │   │
│  │ • Users.js → User CRUD + profiles               │   │
│  │ • Projects.js → Project lifecycle               │   │
│  │ • Tasks.js → Task operations + status flow      │   │
│  │ • Comments.js → Comments + mention extraction   │   │
│  │ • Notifications.js → 5 email notification types │   │
│  │ • Settings.js → User preference persistence     │   │
│  └──────────────────────────────────────────────────┘   │
│                                                          │
│  ┌──────────────────────────────────────────────────┐   │
│  │ Data Access Layer (DataStore.js)                 │   │
│  │ • Row-level CRUD (append, update, delete)       │   │
│  │ • ID generation with auto-increment locking     │   │
│  │ • 3-level caching strategy                      │   │
│  │ • Cache invalidation on mutations               │   │
│  └──────────────────────────────────────────────────┘   │
│                                                          │
│  ┌──────────────────────────────────────────────────┐   │
│  │ Utilities Layer (Utilities.js)                   │   │
│  │ • Email validation, date parsing                │   │
│  │ • Header mapping, safe mail sending             │   │
│  └──────────────────────────────────────────────────┘   │
└────────────────────┬────────────────────────────────────┘
                     │ Google Sheets API
┌────────────────────┴────────────────────────────────────┐
│        Google Sheets (Data Persistence Layer)           │
│                                                          │
│  Tabs: Users | Projects | Tasks | Comments | Assignments
│  (One tab per entity, human-readable ID columns)        │
└──────────────────────────────────────────────────────────┘
```

### Frontend Architecture (Vue 3)

- **Reactive Data**: Refs for `users`, `projects`, `tasks`, `assignments`, `comments`
- **State Management**: Computed properties for derived state (current user, assignee summaries)
- **Component Structure**: Modular views (Kanban, All Items, Modals) with reusable logic
- **Data Sync**: Version hashing to detect stale data; memoization to prevent unnecessary re-renders
- **UI Framework**: Tailwind CSS for responsive design, Font Awesome icons

### Backend Communication

- **Web App Access**: `ANYONE` access, executes as `USER_ACCESSING`
- **Request/Response**: Functions return JSON strings for client parsing
- **Error Handling**: All responses include success flag and error messages
- **Locking**: `LockService` prevents race conditions during ID generation

### Caching Strategy (3 Levels)

| Level | Scope | TTL | Purpose |
|-------|-------|-----|---------|
| **REQUEST_CACHE** | Per request | ~500ms | In-memory storage during single function call |
| **CacheService** | Script-level | 300s (5 min) | Users table cache (most frequently accessed) |
| **ScriptProperties** | Persistent | ∞ | ID counters, user settings, data version |

**Invalidation**: Every write operation bumps `soar_data_version:last_updated` timestamp, forcing clients to refresh.

### Permission Model

- **Manager Hierarchy**: Users have optional `Manager_ID` (creates reporting chain)
- **Assignment Permissions**: Can assign tasks only to:
  - Self
  - Direct subordinates (where `Manager_ID = current_user`)
- **Visibility**: All users can see projects/tasks assigned to them
- **Admin**: Project creator can always modify project settings

### Data Flow Example: Creating a Task

```
1. User clicks "New Task" in Kanban → Vue modal opens
2. User fills form (title, description, priority, assignees, due date)
3. User clicks "Create" → Vue calls GAS function createTask()
4. GAS validates inputs (required fields, valid assignee IDs, date format)
5. GAS generates next Task_ID (T-00000042) with LockService
6. GAS appends row to Tasks sheet
7. GAS appends Assignment rows (one per assignee)
8. GAS invalidates cache (bumps data_version timestamp)
9. GAS sends notification emails to all assignees (async via Notifications.js)
10. GAS returns JSON with new task object
11. Vue updates local state and re-renders Kanban board
12. Other browsers detect version change on next poll, refresh data
```

---

## Data Model

### Entity Relationships

```
User (U-00000001)
├── manages → User[] (recursion via Manager_ID)
├── creates → Project[] (Projects.Creator_ID)
├── completes → Task[] (Tasks.Completed_By)
└── comments on → Comment[] (Comments.Commenter_ID)

Project (P-00000001)
├── contains → Task[] (Tasks.Project_ID)
├── assigned to → User[] (via Assignments table)
└── receives → Comment[] (Comments.Topic_ID)

Task (T-00000001)
├── belongs to → Project (Tasks.Project_ID)
├── assigned to → User[] (via Assignments table)
├── receives → Comment[] (Comments.Topic_ID)
└── completed by → User (Tasks.Completed_By, optional)

Comment (C-00000001)
├── on → Task or Project (Comments.Topic_ID)
├── by → User (Comments.Commenter_ID)
└── mentions → User[] (extracted from content)

Assignment (Task_ID or Project_ID)
├── connects → Task / Project (Assignment_ID)
└── to → User (Assignments.Assignee_ID)
```

### ID Generation Pattern

All primary keys are human-readable, auto-incrementing, zero-padded:
- **Users**: `U-00000001`, `U-00000002`, ... `U-99999999`
- **Projects**: `P-00000001`, `P-00000002`, ... `P-99999999`
- **Tasks**: `T-00000001`, `T-00000002`, ... `T-99999999`
- **Comments**: `C-00000001`, `C-00000002`, ... `C-99999999`

Generated by Apps Script with synchronized locking to prevent race conditions.

### Data Dictionary

> **Notes**
> - **"Can Be Null = No"** means the field is required at creation time unless auto-filled by the system.
> - **Foreign Keys**: References implemented at application layer (not database foreign keys).

#### Users

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `User_ID` | String, auto-increment | Format: `U-00000000` | No |
| `Email` | Email address | Unique key, used for user login | No |
| `Name` | String | User display name | No |
| `Manager_ID` | String, User_ID reference | Manager's User_ID (for permissions) | Yes |
| `Profile_Pic_Url` | URL | Google Account profile photo | Yes |

#### Projects

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Project_ID` | String, auto-increment | Format: `P-00000000` | No |
| `Project_Title` | String | Display title | No |
| `Description` | String | Goals and scope | Yes |
| `Status` | String | One of: `Not Started`, `In Progress`, `Completed`, `Delayed` | No |
| `Created_Date` | Date | Auto-populated at creation | No |
| `Due_Date` | Date | Planned completion date | Yes |
| `Creator_ID` | String, User_ID reference | Project creator | No |

#### Tasks

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Task_ID` | String, auto-increment | Format: `T-00000000` | No |
| `Project_ID` | String | Parent project (Projects.Project_ID) | No |
| `Task_Title` | String | Display title | No |
| `Description` | String | Task details | Yes |
| `Status` | String | One of: `Not Started`, `In Progress`, `Completed`, `Delayed` | No |
| `Priority` | String | One of: `High`, `Medium`, `Low` | Yes |
| `Created_Date` | Date | Auto-populated at creation | No |
| `Due_Date` | Date | Planned completion date | Yes |
| `Completed_By` | String, User_ID reference | User who completed (Tasks.Completed_By) | Yes |
| `Completed_At` | DateTime | Timestamp of completion | Yes |

#### Comments

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Comment_ID` | String, auto-increment | Format: `C-00000000` | No |
| `Topic_ID` | String | Task_ID or Project_ID (polymorphic) | No |
| `Topic_Type` | String | One of: `Task`, `Project` | No |
| `Commenter_ID` | String, User_ID reference | Comment author | No |
| `Content` | String | Comment text (may contain @mentions) | No |
| `Timestamp` | DateTime | Auto-populated at creation | No |
| `Is_Resolved` | Boolean | Whether comment is marked resolved | No |

#### Assignments

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Assignment_ID` | String | Task_ID or Project_ID | No |
| `Assignee_ID` | String | User_ID of assigned user | No |

---

## Configuration

### Environment Setup

#### appsscript.json

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
- **timeZone**: Used for due date handling and weekly digest scheduling (adjust to your region)
- **oauthScopes**: Permissions for reading sheets, sending emails, fetching profile photos
- **webapp.access**: `ANYONE` allows public access (authentication via Google Account is implicit)
- **webapp.executeAs**: `USER_ACCESSING` runs script with deployer's permissions (recommended for safety)
- **runtimeVersion**: `V8` required (old Rhino runtime not supported)

#### Configuration Constants

In [DataStore.js](DataStore.js), update these if needed:

| Constant | Default | Purpose |
|----------|---------|---------|
| `REQUEST_CACHE` | `{}` | Per-request in-memory cache |
| `USER_LIST_CACHE_SECONDS` | `300` | Cache TTL for Users table (5 min) |
| `DATA_VERSION_PROPERTY_KEY` | `soar_data_version:last_updated` | ScriptProperties key for cache invalidation |
| `ID_COUNTER_PROPERTY_PREFIX` | `soar_next_id:` | Prefix for auto-increment counters |
| `USER_SETTINGS_PROPERTY_PREFIX` | `soar_user_settings:` | Prefix for per-user settings |

#### Spreadsheet Setup

Each sheet tab requires specific column headers (exact order matters):

**Users**:
```
User_ID | Email | Name | Manager_ID | Profile_Pic_Url
```

**Projects**:
```
Project_ID | Project_Title | Description | Status | Created_Date | Due_Date | Creator_ID
```

**Tasks**:
```
Task_ID | Project_ID | Task_Title | Description | Status | Priority | Created_Date | Due_Date | Completed_By | Completed_At
```

**Comments**:
```
Comment_ID | Topic_ID | Topic_Type | Commenter_ID | Content | Timestamp | Is_Resolved
```

**Assignments**:
```
Assignment_ID | Assignee_ID
```

---

## API Reference

### User Functions

#### `addUser(email, name, managerEmail)`
Creates a new user and returns JSON response.

**Parameters**:
- `email` (string): User email (must be unique)
- `name` (string): User display name
- `managerEmail` (string, optional): Email of user's manager

**Returns**: `{success: true, user_id: "U-00000001"}` or `{success: false, error: "..."}`

**Called by**: Bootstrap on first app load, Settings modal

---

#### `updateCurrentUserProfile(name, managerEmail)`
Updates logged-in user's name and manager.

**Parameters**:
- `name` (string): New display name
- `managerEmail` (string, optional): New manager email

**Returns**: `{success: true}` or `{success: false, error: "..."}`

---

### Project Functions

#### `createProject(title, description, status, dueDate)`
Creates a new project.

**Parameters**:
- `title` (string): Project title
- `description` (string, optional): Project description
- `status` (string): One of `Not Started`, `In Progress`, `Completed`, `Delayed`
- `dueDate` (string, optional): Date in `YYYY-MM-DD` format

**Returns**: `{success: true, project_id: "P-00000001", project: {...}}`

---

#### `updateProject(projectId, title, description, status, dueDate)`
Updates project metadata.

**Parameters**:
- `projectId` (string): Project_ID to update
- `title` (string): New title
- `description` (string): New description
- `status` (string): New status
- `dueDate` (string, optional): New due date

**Returns**: `{success: true, project: {...}}`

---

#### `deleteProject(projectId)`
Deletes a project and all cascading tasks/assignments.

**Parameters**:
- `projectId` (string): Project_ID to delete

**Returns**: `{success: true}`

---

### Task Functions

#### `createTask(projectId, title, description, priority, dueDate, assigneeEmails)`
Creates a new task with optional assignees.

**Parameters**:
- `projectId` (string): Parent Project_ID
- `title` (string): Task title
- `description` (string, optional): Task description
- `priority` (string, optional): One of `High`, `Medium`, `Low`
- `dueDate` (string, optional): Date in `YYYY-MM-DD` format
- `assigneeEmails` (array, optional): List of email addresses to assign

**Returns**: `{success: true, task_id: "T-00000001", task: {...}}`

---

#### `updateTask(taskId, title, description, status, priority, dueDate, assigneeEmails)`
Updates task metadata and assignees.

**Parameters**:
- `taskId` (string): Task_ID to update
- `title` (string): New title
- `description` (string): New description
- `status` (string): New status (triggers completion logic if `Completed`)
- `priority` (string): New priority
- `dueDate` (string, optional): New due date
- `assigneeEmails` (array): List of assignee emails (replaces existing)

**Returns**: `{success: true, task: {...}}`

---

#### `completeTask(taskId)`
Marks a task as completed and records completion timestamp and completing user.

**Parameters**:
- `taskId` (string): Task_ID to complete

**Returns**: `{success: true}`

**Net Effect**: Sets `Status = "Completed"`, `Completed_By = current_user`, `Completed_At = now()`

---

#### `deleteTask(taskId)`
Deletes a task and cascading assignments.

**Parameters**:
- `taskId` (string): Task_ID to delete

**Returns**: `{success: true}`

---

### Comment Functions

#### `addComment(topicId, topicType, content)`
Creates a new comment with automatic mention extraction.

**Parameters**:
- `topicId` (string): Task_ID or Project_ID to comment on
- `topicType` (string): Either `Task` or `Project`
- `content` (string): Comment text (can include @mention usernames)

**Returns**: `{success: true, comment_id: "C-00000001"}`

**Side Effects**: Extracts mentions from `content` and triggers notification emails.

---

#### `deleteComment(commentId)`
Deletes a comment.

**Parameters**:
- `commentId` (string): Comment_ID to delete

**Returns**: `{success: true}`

---

#### `resolveComment(commentId)`
Marks a comment as resolved (does not delete).

**Parameters**:
- `commentId` (string): Comment_ID to resolve

**Returns**: `{success: true}`

---

### Notification Functions

#### `notifyAssignment(userEmail, taskId, taskTitle)`
Sends email when user is assigned to a task (internal use).

---

#### `notifyMention(mentionedEmail, commenterName, topicId, topicType)`
Sends email when user is mentioned in a comment (internal use).

---

#### `sendWeeklyDigest()`
Sends weekly summary of assigned tasks to all users (scheduled).

---

### Bootstrap & Initialization

#### `getInitialPayload()`
Called on app load to fetch all user data for cache warming.

**Returns**: JSON object containing:
```json
{
  "current_user": {...},
  "users": [...],
  "projects": [...],
  "tasks": [...],
  "assignments": [...],
  "comments": [...],
  "settings": {...}
}
```

---

## Development

### Adding a New Feature

#### 1. Extend the Data Model

If your feature requires new data:
1. Create a new sheet tab for the entity
2. Define the column headers (update [Configuration](#configuration) section)
3. Add CRUD functions to a new service file (e.g., `YourFeature.js`)

#### 2. Add Backend Logic

Create a new file (e.g., `YourFeature.js`) following the service pattern:

```javascript
function createNewThing(name, description) {
  // Input validation
  if (!name || name.trim() === '') {
    return {success: false, error: 'Name is required'};
  }
  
  // Generate ID
  const newId = DataStore.getNextId('YourFeature');
  
  // Write to sheet
  const sheet = DataStore.getOrCreateSheet('YourFeature');
  sheet.appendRow([
    newId,
    name,
    description,
    new Date()  // created_date auto-fill
  ]);
  
  // Invalidate cache
  DataStore.invalidateTableCache('YourFeature');
  
  // Return success response
  return {success: true, thing_id: newId};
}
```

#### 3. Extend Frontend UI

In [App.js.html](App.js.html), add Vue components as needed:
- New data refs in the initial state
- Fetch functions to call your backend GAS functions
- New view sections or modals
- Update navigation to expose the feature

#### 4. Add Email Notifications (if applicable)

In [Notifications.js](Notifications.js), add a new notification function:

```javascript
function notifyYourEvent(userEmail, eventDetails) {
  if (!getUserSetting(userEmail, 'notify_your_event')) {
    return;  // User opted out
  }
  
  MailApp.sendEmail(
    userEmail,
    'Subject',
    'Email body'
  );
}
```

#### 5. Test & Document

- Test in the Apps Script editor using `Test Deployments`
- Verify frontend calls your new function correctly
- Document your API in the [API Reference](#api-reference) section
- Update [Data Model](#data-model) if you added entities

### Code Organization

- **Utilities.js**: Shared helpers (validation, formatting)
- **DataStore.js**: Data access layer (all sheet I/O happens here)
- **[Feature].js**: Business logic services (one file per domain feature)
- **Code.js**: Web app entry point (`doGet()`)
- **Bootstrap.js**: Initial data loading
- **App.js.html**: Vue.js frontend (UI rendering + client-side state)

### Running Tests

Currently, Apps Script testing is manual:
1. Deploy as test deployment: **Deploy** → **Test deployments**
2. Open the URL in browser
3. Verify features work end-to-end
4. Check browser console for errors
5. Review Apps Script logs: **Executions** tab

### Code Style & Conventions

- **Functions**: camelCase (e.g., `createTask()`, `getInitialPayload()`)
- **Variables**: camelCase for local, UPPER_CASE for constants
- **Error Handling**: Always return `{success: false, error: "message"}` on failure
- **Comments**: Use JSDoc-style for exported functions
- **Validation**: Validate all inputs at function entry point
- **Caching**: Always call `DataStore.invalidateTableCache(tableName)` after writes

---

## Troubleshooting

### App won't load / Blank screen

**Causes & Solutions**:
1. **Deployment URL is wrong**: Verify you deployed as "Web app" (not just saving), and access the correct URL
2. **Apps Script hasn't mounted Vue**: Check browser console for JavaScript errors
3. **Sheet tabs missing**: Verify all 5 sheet tabs exist with correct names: `Users`, `Projects`, `Tasks`, `Comments`, `Assignments`
4. **Apps Script permissions not granted**: Refresh page and re-authenticate

**Debug steps**:
- Open browser DevTools (F12)
- Check **Console** tab for errors
- Check **Network** tab for failed requests
- Open Apps Script editor and check **Executions** log for runtime errors

---

### Tasks/Projects not appearing

**Possible Causes**:
1. **Data not saved yet**: Click "Create Project" and wait—might take a few seconds
2. **Sheet columns are wrong order**: Verify exact column headers match [Configuration](#configuration)
3. **Cache is stale**: On rare occasions, version hash polling may be delayed—refresh browser

**Solution**:
- Check the Google Sheet directly: open sheet tabs and verify data is there
- Check Apps Script logs for errors: **Executions** tab
- Refresh browser to force data reload

---

### Email notifications not sending

**Possible Causes**:
1. **Gmail API not authorized**: Script needs permission—accept the first time
2. **User has notifications disabled**: Check [Settings](#accessibility--personalization) modal for this notification type
3. **Email invalid**: Verify recipient email address is in correct format

**Debug**:
- Check Apps Script **Execution log** for email send errors
- Verify sender email address is authorized Gmail account
- Test manually: Call `notifyAssignment()` directly in Apps Script editor

---

### "You do not have permission to access this file"

**Cause**: You're trying to access a deployment URL that you don't own or Apps Script project was deleted.

**Solution**:
1. Verify you have edit access to the Google Sheet
2. Go to **Extensions** → **Apps Script** to open the correct project
3. Deploy a fresh web app: **Deploy** → **New deployment**
4. Use the new URL

---

### Slow performance / High latency

**Causes**:
1. **Google Sheets is slow**: Large sheets (1000+ rows) slow down queries
2. **Too many concurrent users**: Apps Script has quota limits (~1000 reads/sec across all projects)
3. **Cache invalidation storm**: Rapid mutations clear cache, forcing full reloads

**Solutions**:
1. Archive old completed projects/tasks (move to separate "Archive" sheet)
2. Batch operations when possible (multiple updates in one function call)
3. Review [Caching Strategy](#caching-strategy--3-levels) to tune TTLs

---

### Mention extraction not working in comments

**Problem**: `@username` mentions are not being recognized / notified.

**Causes**:
1. **Typo in username**: The @mention must match exact display name (case-sensitive)
2. **User not in Users table yet**: Only onboarded users can be mentioned
3. **Regex bug in mention extraction**: Unlikely, but check [Comments.js](Comments.js)

**Solution**:
- Copy exact username (with spaces if present): `@John Smith` not `@johnsmith`
- Ensure user is already in the system (check Users sheet)
- Manually test mention regex: Open Apps Script editor, paste code in console

---

### Permissions error: "Cannot assign task to this user"

**Cause**: You're trying to assign a task to someone outside your reporting chain.

**Why**: Permission model restricts assignments to:
- Yourself
- Your direct subordinates (users where `Manager_ID = your_user_id`)

**Solution**:
- Assign tasks to teammates in your group
- Have manager create cross-team task and assign to you
- Contact system admin to adjust manager hierarchy in Users sheet

---

### "Generated IDs reset / duplicated IDs"

**Cause**: ScriptProperties counter corrupted (rare).

**Workaround**:
1. Open Apps Script editor
2. Go to **Services** → **Execution log**
3. Run this in the console:
   ```javascript
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Users');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Projects');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Tasks');
   PropertiesService.getScriptProperties().deleteProperty('soar_next_id:Comments');
   ```
4. Refresh the app—counters will re-seed from sheet data on next mutation

---

## Support & Contributing

### Reporting Issues

Found a bug? Have a feature request?
1. Check [Troubleshooting](#troubleshooting) section first
2. Open an issue on [GitHub](https://github.com/nathanwiggins/soar/issues)
3. Include: steps to reproduce, expected vs actual, browser/OS, error logs

### Contributing

We welcome contributions! To add a feature or fix a bug:
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make changes following [Code Style](#code-style--conventions)
4. Test thoroughly (end-to-end in deployed app)
5. **Update README** if you add features or change behavior
6. Submit a pull request with description of changes

### Roadmap

Potential future enhancements:
- [ ] File attachments on tasks/comments
- [ ] Recurring tasks
- [ ] Time tracking / Kanban burn-down charts
- [ ] Mobile app (React Native)
- [ ] Slack integration for notifications
- [ ] Calendar view of due dates
- [ ] Advanced filters and saved views
- [ ] Bulk import from CSV

---

## License

Soar is provided as-is for educational and professional use within Google Workspace environments.

---

## Credits

Built with ❤️ using Google Apps Script, Vue.js, and Tailwind CSS.
# Soar

## 1. Backend and Data Dictionary

This project management system is implemented using **Google Apps Script** as the backend logic layer and **Google Sheets** as the primary data store. Apps Script handles data validation, business rules, ID generation, timestamps, and CRUD workflows, while Sheets stores each logical entity in a dedicated tab.

### 1.1 Backend Architecture (Apps Script + Google Sheets)

- **Platform:** Google Apps Script (JavaScript runtime in Google Workspace).
- **Database Layer:** Google Sheets workbook acting as a relational-style datastore.
- **Entity Storage Pattern:** One sheet tab per table (`Users`, `Projects`, `Tasks`, `Comments`, `Assignments`).
- **Record Identity Pattern:** Human-readable prefixed IDs with zero padding:
  - Users: `U-00000000`
  - Projects: `P-00000000`
  - Tasks: `T-00000000`
  - Comments: `C-00000000`
- **Automation Rules:**
  - Auto-increment IDs generated by Apps Script.
  - `Created_Date` / `Timestamp` auto-populated at insert time.
  - Referential integrity enforced in script logic (e.g., `Creator_ID`, `Project_ID`, `Commenter_ID`, `Assignee_ID` must exist in `Users` or related parent tables).
- **Core Responsibilities in Script Services:**
  - Input validation (email format, URL format, date types, integer ranges).
  - Status normalization (e.g., `Not Started`, `In Progress`, `Completed`, `Delayed`).
  - Nullability enforcement based on schema.
  - Lookup/joins for display views (e.g., tasks with assignee and project metadata).

### 1.2 Data Dictionary

> **Notes**
> - “Can Be Null = No” means the field is required at creation time unless explicitly auto-filled by the system.
> - Fields marked as aliases are foreign-key-like references implemented at the application level.

#### Users

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `User_ID` | Int, auto-increment, format: `U-00000000` | Unique identifier for a user. | No |
| `Email` | Valid email address | User email; unique key used to identify each user. | No |
| `Name` | String | User display name (captured at onboarding/first use). | No |
| `Manager_ID` | Int, `User_ID` alias | The user’s manager (`Users.User_ID`). | Yes |
| `Profile_Pic_Url` | URL | URL to user profile image displayed in the app. | Yes |

#### Projects

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Project_ID` | Int, auto-increment, format: `P-00000000` | Unique identifier for a project. | No |
| `Project_Title` | String | Descriptive title displayed to users. | No |
| `Description` | String | Detailed project description, goals, and expectations. | Yes |
| `Status` | String | Project state (`Not Started`, `In Progress`, `Completed`, `Delayed`, etc.). | No |
| `Created_Date` | Date, auto-fill | Date the project was created. | No |
| `Due_Date` | Date | Planned completion date for the project. | Yes |
| `Creator_ID` | Int, `User_ID` alias | User ID of project creator. | No |

#### Tasks

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Task_ID` | Int, auto-increment, format: `T-00000000` | Unique identifier for a task. | No |
| `Project_ID` | Int | Identifier of the parent project. | No |
| `Task_Title` | String | Task title displayed to users. | No |
| `Due_Date` | Date | Planned completion date for the task. | Yes |
| `Status` | String | Task state (`Not Started`, `In Progress`, `Completed`, `Delayed`, etc.). | No |
| `Created_Date` | Date, auto-fill | Date the task was created. | No |
| `Description` | String | Detailed task description. | Yes |
| `Priority` | Enum (`High`, `Medium`, `Low`) | Task urgency level. | Yes |
| `Completed_By` | Int, `User_ID` alias | User ID of the person who completed the task. | Yes |
| `Completed_At` | Date/Time, auto-fill | Date and time the task was marked completed. | Yes |

#### Comments

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Comment_ID` | Int, auto-increment, format: `C-00000000` | Unique identifier for a comment. | No |
| `Topic_ID` | Int, `Task_ID` or `Project_ID` alias | Identifier of the task or project being commented on. | No |
| `Timestamp` | Date/Time, auto-fill | Date and time comment was created. | No |
| `Commenter_ID` | Int, `User_ID` alias | User ID of the commenter. | No |
| `Content` | String | Comment text content. | No |

#### Assignments

| Field | Type | Description | Can Be Null |
|---|---|---|---|
| `Assignment_ID` | Int, `Task_ID` or `Project_ID` alias | Identifier for assigned task/project record. | No |
| `Assignee_ID` | Int, `User_ID` alias | Identifier for assigned user. | No |

### 1.3 Relationship Summary

- One **User** can create many **Projects** (`Projects.Creator_ID -> Users.User_ID`).
- One **Project** can contain many **Tasks** (`Tasks.Project_ID -> Projects.Project_ID`).
- One **User** can author many **Comments** (`Comments.Commenter_ID -> Users.User_ID`).
- One **Task** or **Project** can have many **Comments** (`Comments.Topic_ID`).
- One **Task** must always be assigned to at least one **User** through **Assignments**.
- One **Project** can be assigned to one or more **Users** through **Assignments**.
