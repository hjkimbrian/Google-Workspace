# Gmail Signature Manager

A Google Apps Script web app that lets Google Workspace admins view and update Gmail signatures for individual users or the entire domain — directly in the browser, with a WYSIWYG rich-text editor.

## Features

- **User list** — browse all domain users (paginated, with live search)
- **Load signature** — fetch any user's current Gmail signature into the editor
- **WYSIWYG editor** — rich-text formatting (bold, italic, font, colour, links, images, lists, alignment)
- **Template variables** — insert `{{firstName}}`, `{{email}}`, etc. via one-click chips; each user's Directory profile is fetched at save time to personalise their signature automatically
- **Preview** — render the template with a specific user's real data before committing
- **Update one user** — save the rendered signature for the selected user only
- **Apply to all** — push a personalised, variable-substituted signature to every active user in the domain
- **Secure credentials** — the service account private key is stored in Script Properties, never in source code

## Supported template variables

Insert these placeholders into your signature template. They are replaced with each user's live data from the Admin Directory API at save time.

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

Missing fields are silently replaced with an empty string — no `{{placeholder}}` text will ever appear in a sent signature.

**Example template:**
```
<b>{{fullName}}</b><br>
{{jobTitle}} · {{department}}<br>
{{email}} | {{workPhone}}<br>
{{company}}
```

---

## Prerequisites

- Google Workspace super admin account
- A **Google Cloud project** linked to your Apps Script project (see Step 1)
- Ability to configure **domain-wide delegation** in the Google Admin Console

---

## Step 1 — Create or select a Google Cloud Project

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (or select an existing one).
3. Note the **project number** — you'll need it when linking to Apps Script.

---

## Step 2 — Enable required APIs

In your GCP project, enable these two APIs:

1. **Admin SDK API**
   - Navigation: *APIs & Services → Library → search "Admin SDK API" → Enable*

2. **Gmail API**
   - Navigation: *APIs & Services → Library → search "Gmail API" → Enable*

---

## Step 3 — Create a service account with domain-wide delegation

### 3a — Create the service account

1. Navigate to *IAM & Admin → Service Accounts → Create Service Account*.
2. Give it a descriptive name, e.g. `signature-manager`.
3. Skip the optional role and user access steps — click **Done**.

### 3b — Create and download a JSON key

1. Click the newly created service account → **Keys** tab → **Add Key → Create new key**.
2. Select **JSON** → **Create**.
3. A `.json` file is downloaded — keep it safe. You will paste its contents into Script Properties in Step 6, then **delete the file**.

### 3c — Enable domain-wide delegation

1. In the service account detail page, click **Edit** (pencil icon).
2. Expand *Advanced settings* → check **Enable Google Workspace Domain-wide Delegation**.
3. Click **Save**.
4. Note the **Client ID** shown under the service account (a numeric string).

---

## Step 4 — Authorise OAuth scopes in Google Admin

1. Log into [admin.google.com](https://admin.google.com) as a super admin.
2. Navigate to *Security → Access and data control → API controls → Manage Domain-Wide Delegation*.
3. Click **Add new**.
4. Enter the **Client ID** from Step 3c.
5. Paste the following scopes (comma-separated) into the OAuth scopes field:

```
https://www.googleapis.com/auth/admin.directory.user.readonly,
https://www.googleapis.com/auth/gmail.settings.basic
```

6. Click **Authorise**.

> **Why these scopes?**
> - `admin.directory.user.readonly` — lets the app list users via the Admin Directory API, impersonating your admin account.
> - `gmail.settings.basic` — lets the app read and write Gmail send-as settings (including signatures), impersonating individual users.

---

## Step 5 — Create the Apps Script project

1. Go to [script.google.com](https://script.google.com) → **New project**.
2. Rename the project (e.g. *Gmail Signature Manager*).

### 5a — Link your GCP project

This is required so the script can use the APIs enabled in Step 2.

1. In the Apps Script editor: *Project Settings* (gear icon) → **Change project** under *Google Cloud Platform (GCP) Project*.
2. Enter your GCP **project number** from Step 1 → **Set project**.

### 5b — Create the script files

Create four files in the editor — one for each file in this folder:

| Apps Script file | Source file in this repo         |
|------------------|----------------------------------|
| `Code.gs`        | `SignatureManager/Code.gs`       |
| `Auth.gs`        | `SignatureManager/Auth.gs`       |
| `Index.html`     | `SignatureManager/Index.html`    |

> To create a new file: click the **+** button next to *Files* in the left panel.
> For `.gs` files choose *Script*; for `.html` choose *HTML*.
> Paste the contents of each source file exactly as written.

---

## Step 6 — Store credentials in Script Properties

Script Properties are encrypted at rest and are never visible in source code.

1. In the Apps Script editor: *Project Settings* (gear icon) → scroll to **Script Properties** → **Add script property**.
2. Add the following three properties:

| Property name         | Value                                                        |
|-----------------------|--------------------------------------------------------------|
| `SERVICE_ACCOUNT_KEY` | The **entire contents** of the JSON key file from Step 3b.  |
| `ADMIN_EMAIL`         | Your super admin email, e.g. `admin@example.com`            |
| `DOMAIN`              | Your primary domain, e.g. `example.com`                     |

3. Click **Save script properties**.
4. **Delete the downloaded JSON key file** from your computer — it is now stored securely in Script Properties.

---

## Step 7 — Deploy the web app

1. In the Apps Script editor, click **Deploy → New deployment**.
2. Click the gear icon next to *Select type* → choose **Web app**.
3. Configure:
   - **Description** — e.g. `v1`
   - **Execute as** — `Me (owner)` *(the service account calls are made server-side as the script owner)*
   - **Who has access** — `Anyone within [your domain]` *(restricts access to Workspace accounts in your domain)*
4. Click **Deploy**.
5. Copy the **Web app URL** shown — share this URL with other admins.

> **Re-deploying after changes:** use *Deploy → Manage deployments → Edit* (pencil) → bump the version number → **Deploy**.

---

## Usage

1. Open the web app URL.
2. The left panel lists all users in your domain. Use the search box to filter.
3. Click a user to load their current signature into the editor.
4. Edit the signature using the rich-text toolbar.
5. Click one of:
   - **Update This User** — saves the signature for the selected user only.
   - **Apply to All Users** — pushes the signature to every active user in the domain (a confirmation dialog will appear first).

---

## Troubleshooting

### "SERVICE_ACCOUNT_KEY not set in Script Properties"
You skipped or did not save Step 6. Verify the property name is exactly `SERVICE_ACCOUNT_KEY`.

### "Token exchange failed: unauthorized_client"
Domain-wide delegation is not configured correctly. Double-check:
- The service account has **DWD enabled** (Step 3c).
- The correct **Client ID** and both **OAuth scopes** are authorised in the Admin Console (Step 4).
- The GCP project linked in Apps Script is the same project that owns the service account (Step 5a).

### "Directory API error (403)"
The `ADMIN_EMAIL` account does not have super admin rights, or the Admin SDK API is not enabled (Step 2).

### "Apply to All" times out for large domains
Apps Script has a **6-minute execution limit**. For domains with many hundreds of users, the bulk update may not complete in a single run. Workarounds:
- Run the update during off-peak hours so each API call returns faster.
- Consider splitting users alphabetically across multiple manual runs using the *Update This User* button.
- For very large domains, export the script logic to a Cloud Run job or use the [Gmail API batch endpoint](https://developers.google.com/workspace/gmail/api/guides/batch).

### Signature HTML is simplified after editing
Quill's editor normalises some complex HTML (e.g. inline `style` attributes for custom fonts). If your existing signatures use advanced HTML, consider using the Quill editor for new, standardised signatures and testing the output before applying to all users.

---

## Security notes

- The service account private key is stored **only** in Script Properties — never in source code or version control.
- The web app executes **as the owner** (admin who deployed it), but individual user impersonation happens server-side via domain-wide delegation.
- Access is restricted to **domain users only** via the deployment setting.
- After downloading the JSON key in Step 3b, delete the local file immediately once you've saved it to Script Properties.
