# GWS Admin Toolkit

A Google Apps Script web app for Google Workspace super admins. Provides four tools in a single deployment:

- **Signatures** — view and update Gmail signatures for individual users or the entire domain, with a WYSIWYG rich-text editor and template variables.
- **Data Transfer** — transfer Google Drive files, Calendar data, and Looker Studio reports from one user account to another using the Google Workspace Data Transfer API.
- **Delegation** — manage Gmail delegates and primary calendar sharing for any user, with policy visibility (mail delegation allowed/disallowed per OU via the Cloud Identity Policy API).
- **Licenses** — view, assign, switch, and remove Google Workspace licenses per user; synchronise licenses based on Google Group membership with dry-run preview and optional Google Sheets audit log.

---

## Features

### Signatures

- **User list** — browse all domain users (paginated, with live search and Gmail-only filter)
- **Load signature** — fetch any user's current Gmail signature into the editor
- **WYSIWYG editor** — rich-text formatting (bold, italic, font, colour, links, images, lists, alignment)
- **HTML source mode** — toggle between visual editor and raw HTML textarea
- **Template variables** — insert `{{firstName}}`, `{{email}}`, etc. via one-click chips; each user's Directory profile is fetched at save time to personalise the signature
- **Preview** — render the template with a specific user's real data before saving
- **Update one user** — save the rendered signature for the selected user only
- **Apply to all** — push a personalised signature to every active user in the domain

### Data Transfer

- **Multi-service transfers** — select any combination of Google Drive, Google Calendar, and Looker Studio in a single transfer request
- **Autocomplete user picker** — source and destination inputs use the same domain user list loaded by the Signatures tab
- **Service-specific options** (radio buttons):
  - **Drive**: All files / Private only / Shared only (`PRIVACY_LEVEL` parameter)
  - **Calendar**: Release or keep calendar resources (`RELEASE_RESOURCES` parameter)
  - **Looker Studio**: no additional parameters required
- **Transfer history** — each submitted transfer appears as a card with per-service status badges
- **Refresh status** — poll the Data Transfer API for updated status without leaving the page

### Delegation

- **User picker** — type any domain user to load their delegation and calendar info
- **Policy badge** — resolves the mail delegation policy for the user's OU via the Cloud Identity Policy API; displays Allowed / Not Allowed / Unknown with a visual indicator
- **Mail delegates** — list current delegates with verification status (accepted / pending); add or remove delegates
- **Calendar sharing** — lists current ACL entries for the user's primary calendar; add a new share rule with role radio buttons (Free/Busy, View all details, Make changes); remove existing shares; no email notifications are sent on any ACL change (`sendNotifications=false`)
- **Calendar guard** — the calendar sharing section only appears if the user has Google Calendar enabled

### License Management

- **Product & SKU selector** — choose from the full catalog of Google Workspace, Vault, Drive Storage, Cloud Identity, and Chrome Enterprise products/SKUs
- **Inventory count** — instantly see how many licenses are assigned for the selected SKU across the domain
- **User license view** — pick any user to see all their currently assigned licenses with remove buttons
- **Assign** — add the selected SKU to a user who does not yet hold it
- **Switch** — swap a user's existing license within the same product to the selected SKU
- **Remove** — unassign any individual license from a user
- **412 handling** — if an OU has automatic licensing enabled, the error is surfaced with a clear explanation
- **Group sync** — select a Google Group and target SKU; the sync engine evaluates every member using four rules:
  - `ASSIGN` — member has no license in this product → assign the target SKU
  - `SWITCH` — member has a different SKU in this product → switch to the target SKU
  - `REMOVE` — non-member currently holds the target SKU → revoke it
  - `NO_CHANGE` — member already has the target SKU → skip
- **Dry run mode** — preview all proposed changes (with counts by action type) in a modal before applying anything
- **Google Sheets log** — optionally write sync results to a new or existing Google Sheet; each row records the user, action, before/after SKU, and timestamp

---

## Supported template variables

Insert these placeholders into signature templates. They are replaced with each user's live data from the Admin Directory API at save time.

| Variable | Source field | Example value |
|---|---|---|
| `{{firstName}}` | `name.givenName` | `Jane` |
| `{{lastName}}` | `name.familyName` | `Smith` |
| `{{fullName}}` | `name.fullName` | `Jane Smith` |
| `{{email}}` | `primaryEmail` | `jane@example.com` |
| `{{workPhone}}` | `phones[type=work]` | `+1 415 555 0100` |
| `{{mobilePhone}}` | `phones[type=mobile]` | `+1 415 555 0199` |
| `{{jobTitle}}` | `organizations[primary].title` | `Senior Engineer` |
| `{{department}}` | `organizations[primary].department` | `Engineering` |
| `{{company}}` | `organizations[primary].name` | `Acme Corp` |

Missing fields are silently replaced with an empty string.

---

## Prerequisites

- Google Workspace super admin account
- A **Google Cloud project** linked to your Apps Script project (see Step 1)
- Ability to configure **domain-wide delegation** in the Google Admin Console

---

## Step 1 — Create or select a Google Cloud Project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (or select an existing one).
3. Note the **project number** — you will need it when linking to Apps Script.

---

## Step 2 — Enable required APIs

### 2a — In the GCP project

1. **Admin SDK API** — covers Directory, Data Transfer, and OU lookups
   - *APIs & Services → Library → search "Admin SDK API" → Enable*

2. **Gmail API**
   - *APIs & Services → Library → search "Gmail API" → Enable*

3. **Google Calendar API** *(Delegation tab — Calendar Sharing)*
   - *APIs & Services → Library → search "Google Calendar API" → Enable*

4. **Cloud Identity API** *(Delegation tab — mail delegation policy check)*
   - *APIs & Services → Library → search "Cloud Identity API" → Enable*

5. **Enterprise License Manager API** *(Licenses tab)*
   - *APIs & Services → Library → search "Enterprise License Manager API" → Enable*

> **Note:** The Google Sheets API is no longer required — sync logs are written using the built-in `SpreadsheetApp` service.

### 2b — In the Apps Script editor

Enable the **Admin Directory** advanced service so the script can call the Directory API as the deploying admin without a service account token:

1. Open the script: *Extensions → Apps Script*.
2. Click **Services** (the `+` icon in the left sidebar).
3. Find **Admin SDK API**, select version `directory_v1`, and click **Add**.

The `appsscript.json` manifest in this repo already declares this service — it is enabled automatically when you deploy with [clasp](https://github.com/google/clasp).

---

## Step 3 — Create a service account with domain-wide delegation

### 3a — Create the service account

1. Navigate to *IAM & Admin → Service Accounts → Create Service Account*.
2. Give it a descriptive name, e.g. `gws-admin-toolkit`.
3. Skip the optional role and user access steps → click **Done**.

### 3b — Create and download a JSON key

1. Click the newly created service account → **Keys** tab → **Add Key → Create new key**.
2. Select **JSON** → **Create**.
3. A `.json` file is downloaded — keep it safe. You will paste its entire contents into Script Properties in Step 6, then **delete the local file**.

### 3c — Enable domain-wide delegation

1. In the service account detail page, click **Edit** (pencil icon).
2. Expand *Advanced settings* → check **Enable Google Workspace Domain-wide Delegation**.
3. Click **Save**.
4. Note the **Client ID** (a numeric string) shown under the service account.

---

## Step 4 — Authorise OAuth scopes in Google Admin

1. Log into [admin.google.com](https://admin.google.com) as a super admin.
2. Navigate to *Security → Access and data control → API controls → Manage Domain-Wide Delegation*.
3. Click **Add new**.
4. Enter the **Client ID** from Step 3c.
5. Paste the following scopes (comma-separated):

```
https://www.googleapis.com/auth/gmail.settings.basic,
https://www.googleapis.com/auth/admin.datatransfer,
https://www.googleapis.com/auth/gmail.settings.sharing,
https://www.googleapis.com/auth/calendar,
https://www.googleapis.com/auth/cloud-identity.policies,
https://www.googleapis.com/auth/apps.licensing
```

6. Click **Authorise**.

> **Why these scopes?**
>
> These are the scopes granted to the **service account** for operations that require impersonating a specific user (DWD) or that have no Apps Script advanced service.
>
> | Scope | Used for |
> |---|---|
> | `gmail.settings.basic` | Read and write Gmail send-as settings (signatures) — impersonates each user |
> | `admin.datatransfer` | Create and monitor data transfers — called as admin |
> | `gmail.settings.sharing` | Read, add, and remove Gmail delegates — impersonates each user |
> | `calendar` | Read primary calendar status and manage ACL rules — impersonates each user |
> | `cloud-identity.policies` | Read mail delegation policy per OU (no Apps Script advanced service) |
> | `apps.licensing` | Assign, switch, and remove Google Workspace licenses (no Apps Script advanced service) |
>
> The following are **not** in the DWD list because they are handled by the `AdminDirectory` advanced service and `SpreadsheetApp` built-in, which run as the deploying admin's own OAuth session:
>
> | Removed scope | Now handled by |
> |---|---|
> | `admin.directory.user.readonly` | `AdminDirectory` advanced service |
> | `admin.directory.orgunit.readonly` | `AdminDirectory` advanced service |
> | `admin.directory.group.readonly` | `AdminDirectory` advanced service |
> | `spreadsheets` | `SpreadsheetApp` built-in service |

---

## Step 5 — Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**.
2. Rename the project, e.g. *GWS Admin Toolkit*.

### 5a — Link your GCP project

1. In the Apps Script editor: *Project Settings* (gear icon) → **Change project** under *Google Cloud Platform (GCP) Project*.
2. Enter your GCP **project number** from Step 1 → **Set project**.

### 5b — Create the script files

Create the following files in the editor:

| Apps Script file | Source file in this repo | Notes |
|------------------|--------------------------|-------|
| `Code.gs` | `GWSAdminToolkit/Code.gs` | Main server-side logic |
| `Auth.gs` | `GWSAdminToolkit/Auth.gs` | Service account JWT helper |
| `Index.html` | `GWSAdminToolkit/Index.html` | SPA frontend |
| `Tests.gs` | `GWSAdminToolkit/Tests.gs` | Test suite (optional but recommended) |

> To create a new file: click the **+** button next to *Files*.
> For `.gs` files choose *Script*; for `.html` choose *HTML*.
> Paste the contents of each source file exactly as written.
>
> If you are using [clasp](https://github.com/google/clasp) to sync from this repo,
> the `appsscript.json` manifest is included and will configure the `AdminDirectory`
> advanced service automatically on push.

---

## Step 6 — Store credentials in Script Properties

Script Properties are encrypted at rest and never visible in source code.

1. In the Apps Script editor: *Project Settings* (gear icon) → scroll to **Script Properties** → **Add script property**.
2. Add the following three properties:

| Property name | Value |
|---|---|
| `SERVICE_ACCOUNT_KEY` | The **entire contents** of the JSON key file from Step 3b. |
| `ADMIN_EMAIL` | Your super admin email, e.g. `admin@example.com` |
| `DOMAIN` | Your primary domain, e.g. `example.com` |

3. Click **Save script properties**.
4. **Delete the downloaded JSON key file** from your computer — it is now stored securely in Script Properties.

---

## Step 7 — Deploy the web app

1. In the Apps Script editor, click **Deploy → New deployment**.
2. Click the gear icon next to *Select type* → choose **Web app**.
3. Configure:
   - **Description** — e.g. `v1`
   - **Execute as** — `Me (owner)`
   - **Who has access** — `Anyone within [your domain]`
4. Click **Deploy**.
5. Copy the **Web app URL** — share this with other admins in your domain.

> **Re-deploying after changes:** use *Deploy → Manage deployments → Edit* (pencil) → bump the version number → **Deploy**.

---

## Usage

### Signatures tab

1. Open the web app URL and click the **Signatures** tab (active by default).
2. Use the left panel to browse and search domain users.
3. Click a user to load their current signature into the editor.
4. Edit the signature using the rich-text toolbar, or toggle to raw HTML mode with `</>`.
5. Use the **Insert variable** chips to add personalisation tokens.
6. Click:
   - **Update This User** — saves for the selected user only.
   - **Apply to All Users** — pushes the signature to every active Gmail user in the domain (confirmation dialog first).
   - **Preview** — renders the template with the selected user's real Directory data.

### Data Transfer tab

1. Click the **Data Transfer** tab in the navigation.
2. **From** — type or select the source user's email address.
3. **To** — type or select the destination user's email address.
4. **Service** — click one or more service buttons (Drive, Calendar, Looker Studio). Multiple services can be selected for a single transfer request.
5. **Options** — configure service-specific settings as needed:
   - Drive: choose which files to transfer (all, private only, or shared only).
   - Calendar: choose whether to release calendar resources.
6. Click **Create Transfer**.
7. The transfer appears in the **Transfer History** section with per-service status badges.
8. Click **Refresh** on any card to poll the API for updated status.

> Data transfers are processed asynchronously by Google. Status will show `inProgress` immediately after creation and update to `completed` once finished (typically within minutes to hours depending on data volume).

### Delegation tab

1. Click the **Delegation** tab in the navigation.
2. Type or select a domain user and click **Load**.
3. The **policy badge** shows whether mail delegation is allowed for the user's OU.
   - A warning banner appears if delegation is not allowed for that OU.
4. **Mail Delegation** section:
   - Existing delegates are listed with their verification status.
   - Enter a delegate email address and click **Add Delegate** to grant access.
   - Click the trash icon next to any delegate to remove them.
5. **Calendar Sharing** section (only visible when the user has Calendar enabled):
   - Existing shares are listed with their role badge.
   - Enter a grantee email, choose a role (Free/Busy, View all details, Make changes), and click **Share**.
   - Click the trash icon to remove a calendar share.
   - No email notification is sent to the grantee when a share is added or removed.

### Licenses tab

1. Click the **Licenses** tab in the navigation.
2. Select a **Product** from the dropdown, then select a **SKU** — the count of currently assigned licenses appears immediately.

**User licenses:**

3. Type or select a domain user in the user picker, then click **Load**.
4. All licenses currently assigned to that user are listed with remove buttons.
5. To **assign** the selected SKU, click **Assign**.
6. To **switch** the user from their current SKU (within the same product) to the selected SKU, click **Switch to this SKU**.
7. To **remove** any individual license, click the trash icon next to it.

**Group license sync:**

8. Select a **Group** from the group dropdown.
9. Optionally enable **Write sync log to a Google Sheet** and paste an existing sheet URL (leave blank to create a new sheet in the admin's Drive).
10. Click **Dry Run (preview changes)** to open the preview modal showing every user and the action that would be taken (Assign / Switch / Remove / No change) — no changes are made yet.
11. Click **Apply Changes** inside the dry-run modal to execute, or close and click **Sync Now** to skip the preview and apply with a confirmation dialog.
12. If logging was enabled, a toast appears with the Google Sheet URL once the log is written.

> Group sync rules: members without the target SKU are **assigned** it; members with a different SKU in the same product are **switched**; non-members who hold the target SKU have it **removed**; members already on the correct SKU are skipped (**no change**).

---

## Running the tests

`Tests.gs` contains a suite of unit tests (pure-function, no network) and integration tests (live API calls).

1. Open the Apps Script editor.
2. In the **Functions** dropdown at the top, select `runAllTests`.
3. Click **Run**.
4. Open **Executions** (or press `Ctrl+Enter`) to see the log — each test prints `PASS` or `FAIL` with a reason.

**Unit tests** (always runnable, no setup required):
- `mergeObj_` edge cases
- `substituteVariables_` — placeholder substitution and missing-key behaviour
- `extractVariables_` — Directory profile → template variable mapping, primary org selection
- `getLicenseProducts` — static catalog structure and completeness

**Integration tests** (require valid Script Properties + API access):
- `getConfig` — Script Properties present and well-formed
- `getUsers` — returns users with expected fields; pagination token shape
- `getUserProfile` — fetches name for first domain user
- `getUserId` — resolves email → immutable ID
- `getGroups` — returns array of groups with email field
- `getGroupMembers` — error propagation for invalid group email
- `getLicenseCount` — returns a non-negative number for first catalog SKU
- `syncLicensesForGroup` (dry run) — change list has valid action values; `applied` is `false`

Integration tests are skipped gracefully when optional resources (e.g. no groups in the domain) do not exist.

---

## Troubleshooting

### "SERVICE_ACCOUNT_KEY not set in Script Properties"
You skipped or did not save Step 6. Verify the property name is exactly `SERVICE_ACCOUNT_KEY`.

### "Token exchange failed: unauthorized_client"
Domain-wide delegation is not configured correctly. Check:
- The service account has **DWD enabled** (Step 3c).
- The correct **Client ID** and all three **OAuth scopes** are authorised in the Admin Console (Step 4).
- The GCP project linked in Apps Script is the same project that owns the service account (Step 5a).

### "Directory API error (403)"
The `ADMIN_EMAIL` does not have super admin rights, or the Admin SDK API is not enabled (Step 2).

### "Data Transfer applications API error (403)"
The `admin.datatransfer` scope is missing from DWD authorisation (Step 4), or the Admin Data Transfer API is not enabled on the GCP project (Step 2).

### "Transfer creation failed" — service app ID not found
The Data Transfer page loads available applications from the API when you first visit the tab. If a service button shows no app ID in the warning, ensure the Data Transfer API is enabled and the `admin.datatransfer` DWD scope is authorised, then reload the page.

### "Apply to All" times out for large domains
Apps Script has a **6-minute execution limit**. For domains with many hundreds of users, the bulk signature update may not complete in a single run. Workarounds:
- Run during off-peak hours so each API call returns faster.
- Split users across multiple manual runs using **Update This User**.
- For very large domains, consider exporting the logic to a Cloud Run job or using the [Gmail API batch endpoint](https://developers.google.com/workspace/gmail/api/guides/batch).

### "Failed to add delegate (400): delegation not allowed"
The user's OU has the mail delegation policy set to **Not Allowed**. The policy badge on the Delegation tab will show this. Enable mail delegation for the OU in the Admin Console or Google Workspace Admin policy settings, or move the user to an OU where it is permitted.

### "Policy status: Unknown" on the Delegation tab
Either the `admin.directory.orgunit.readonly` or `cloud-identity.policies` scope is missing from DWD authorisation (Step 4), or the Cloud Identity API is not enabled on the GCP project (Step 2).

### Calendar Sharing section not visible
The selected user does not have Google Calendar enabled. The section is intentionally hidden — check the user's Google Workspace licence and ensure Google Calendar service is enabled for their OU in the Admin Console.

### Signature HTML is simplified after WYSIWYG editing
Quill normalises some complex HTML. Use the `</>` HTML source mode to paste or edit signatures that rely on advanced inline styles.

### "Failed to load license products" or license API errors (403)
The `apps.licensing` scope is missing from DWD authorisation (Step 4), or the Enterprise License Manager API is not enabled on the GCP project (Step 2).

### "Failed to assign license: 412 Precondition Failed"
The user's organisational unit has **automatic licensing** enabled. Google blocks manual license assignment when auto-licensing is active. Disable auto-licensing for the OU in the Admin Console (*Billing → Subscriptions → [product] → Assign licenses automatically*) or move the user to an OU without auto-licensing, then retry.

### "No existing license found in this product to switch from"
The selected user does not have any license in the chosen product. Use **Assign** instead of **Switch**.

### "Failed to load groups" on the Licenses tab
The `admin.directory.group.readonly` scope is missing from DWD authorisation (Step 4).

### Sync log not written / "Failed to write sync log"
The `spreadsheets` scope is missing from DWD authorisation (Step 4), or the Google Sheets API is not enabled on the GCP project (Step 2).

---

## Security notes

- The service account private key is stored **only** in Script Properties — never in source code or version control.
- The web app executes **as the owner** (the admin who deployed it), with individual user impersonation happening server-side via domain-wide delegation.
- Access is restricted to **domain users only** via the deployment setting (*Anyone within [your domain]*).
- Delete the local `.json` key file immediately after saving it to Script Properties.
