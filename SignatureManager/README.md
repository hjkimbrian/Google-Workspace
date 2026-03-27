# GWS Admin Toolkit

A Google Apps Script web app for Google Workspace super admins. Provides two tools in a single deployment:

- **Signatures** — view and update Gmail signatures for individual users or the entire domain, with a WYSIWYG rich-text editor and template variables.
- **Data Transfer** — transfer Google Drive files, Calendar data, and Looker Studio reports from one user account to another using the Google Workspace Data Transfer API.

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

In your GCP project, enable these APIs:

1. **Admin SDK API**
   - *APIs & Services → Library → search "Admin SDK API" → Enable*

2. **Gmail API**
   - *APIs & Services → Library → search "Gmail API" → Enable*

3. **Admin Data Transfer API** *(required for the Data Transfer tab)*
   - *APIs & Services → Library → search "Admin SDK API"* — the Data Transfer API is part of the Admin SDK and is enabled alongside it. If you see it listed separately as "Data Transfer API", enable it too.

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
https://www.googleapis.com/auth/admin.directory.user.readonly,
https://www.googleapis.com/auth/gmail.settings.basic,
https://www.googleapis.com/auth/admin.datatransfer
```

6. Click **Authorise**.

> **Why these scopes?**
> - `admin.directory.user.readonly` — lists domain users and resolves user IDs for the Data Transfer API.
> - `gmail.settings.basic` — reads and writes Gmail send-as settings (signatures).
> - `admin.datatransfer` — creates and monitors data transfers between users.

---

## Step 5 — Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**.
2. Rename the project, e.g. *GWS Admin Toolkit*.

### 5a — Link your GCP project

1. In the Apps Script editor: *Project Settings* (gear icon) → **Change project** under *Google Cloud Platform (GCP) Project*.
2. Enter your GCP **project number** from Step 1 → **Set project**.

### 5b — Create the script files

Create three files in the editor — one for each file in this folder:

| Apps Script file | Source file in this repo |
|------------------|--------------------------|
| `Code.gs` | `SignatureManager/Code.gs` |
| `Auth.gs` | `SignatureManager/Auth.gs` |
| `Index.html` | `SignatureManager/Index.html` |

> To create a new file: click the **+** button next to *Files*.
> For `.gs` files choose *Script*; for `.html` choose *HTML*.
> Paste the contents of each source file exactly as written.

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

### Signature HTML is simplified after WYSIWYG editing
Quill normalises some complex HTML. Use the `</>` HTML source mode to paste or edit signatures that rely on advanced inline styles.

---

## Security notes

- The service account private key is stored **only** in Script Properties — never in source code or version control.
- The web app executes **as the owner** (the admin who deployed it), with individual user impersonation happening server-side via domain-wide delegation.
- Access is restricted to **domain users only** via the deployment setting (*Anyone within [your domain]*).
- Delete the local `.json` key file immediately after saving it to Script Properties.
