# FormAutomation

Google Apps Script that fires on Google Form submission to provision a **Shared Drive** and three **Google Groups** for a new project, then makes the requestor the manager of each group.

## What it does

| Step | Action |
|------|--------|
| 1 | Creates a Shared Drive named after the submitted project. |
| 2 | Creates three Google Groups: **Managers**, **Contributors**, and **Viewers**. Each group email embeds the Shared Drive ID so it's globally unique. |
| 3 | Adds the form requestor as `OWNER` (full manager) of all three groups. |
| *(bonus)* | Writes the Drive ID and group emails back to the response spreadsheet for easy reference. |

## Prerequisites

- Google Workspace account with admin (or delegated) privileges to create Shared Drives and Google Groups.
- The Apps Script project must have the following **Advanced Services** enabled:
  - **Admin SDK Directory API** (`AdminDirectory`)
  - **Drive API v3** (`Drive`)

## Setup

1. Open the Google Form → **⋮ menu → Script editor** (or create a standalone project and link the form).
2. Paste / upload `FormSubmitHandler.gs` into the editor.
3. Edit the configuration block at the top of the file:

   ```javascript
   var READ_EMAIL_FROM_SESSION = false; // true = pull email from respondent login
   var DOMAIN = "example.com";         // your Workspace domain
   var FIELD_PROJECT_NAME    = "project name";  // exact form field title (case-insensitive)
   var FIELD_REQUESTOR_EMAIL = "your email";    // exact form field title (case-insensitive)
   ```

4. Enable Advanced Services:
   - **Resources → Advanced Google Services**
   - Turn on **Admin SDK** and **Drive API**.
5. Add the trigger:
   - **Triggers (clock icon) → Add Trigger**
   - Function: `onFormSubmit`
   - Event source: **From form** (or **From spreadsheet** if using a linked sheet)
   - Event type: **On form submit**
6. Authorize the script when prompted.

## Form fields

The script looks for two fields by title (case-insensitive):

| Field title | Purpose |
|-------------|---------|
| `Project Name` | Names the Shared Drive and groups |
| `Your Email` | Identifies the requestor to add as group manager |

If you set `READ_EMAIL_FROM_SESSION = true` the email field is not required (the script reads the respondent's logged-in email instead, which requires the form to **collect email addresses via sign-in**).

## Group email format

```
<project-slug>-managers-<driveId>@<domain>
<project-slug>-contributors-<driveId>@<domain>
<project-slug>-viewers-<driveId>@<domain>
```

Example for project "My Cool Project" on domain `acme.com`:

```
my-cool-project-managers-0ABCxyz123@acme.com
my-cool-project-contributors-0ABCxyz123@acme.com
my-cool-project-viewers-0ABCxyz123@acme.com
```
