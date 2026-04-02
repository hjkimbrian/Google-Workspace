/**
 * SignatureManager — Code.gs
 *
 * Main server-side entry point for the Gmail Signature Manager web app.
 * All functions in this file are callable from the frontend via
 * google.script.run.<functionName>().
 *
 * Required Script Properties (Project Settings → Script Properties):
 *   SERVICE_ACCOUNT_KEY  — Full JSON of the service account key file.
 *   ADMIN_EMAIL          — Super admin email used to call the Directory API.
 *   DOMAIN               — Primary domain, e.g. "example.com".
 *
 * See README.md for full setup instructions.
 */

// ─── OAuth2 scopes ────────────────────────────────────────────────────────────
//
// These scopes are passed to getServiceAccountToken_() for APIs that require
// DWD user impersonation or have no Apps Script advanced service.
//
// Admin Directory (users, groups, OUs) and Google Sheets are handled by the
// AdminDirectory advanced service and SpreadsheetApp respectively — those
// APIs no longer need service account tokens.

/** Scope for reading and writing Gmail send-as settings (signatures). */
const SCOPE_GMAIL_SETTINGS_ =
  'https://www.googleapis.com/auth/gmail.settings.basic';

/** Scope for the Enterprise License Manager API. */
const SCOPE_LICENSING_ =
  'https://www.googleapis.com/auth/apps.licensing';

// ─── Web app entry point ──────────────────────────────────────────────────────

/**
 * Serves the web app HTML when the deployment URL is visited.
 *
 * Deploy settings (Extensions → Apps Script → Deploy → New deployment):
 *   Execute as: Me (owner)
 *   Who has access: Anyone within [your domain]
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GWS Admin Toolkit')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── Configuration ────────────────────────────────────────────────────────────

/**
 * Returns non-sensitive configuration values to the frontend.
 * Validates that all required Script Properties are present.
 *
 * @returns {{ adminEmail: string, domain: string }}
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const adminEmail = props.getProperty('ADMIN_EMAIL');
  const domain     = props.getProperty('DOMAIN');

  if (!adminEmail || !domain) {
    throw new Error(
      'Missing configuration. Set ADMIN_EMAIL and DOMAIN in ' +
      'Project Settings → Script Properties (see README.md).'
    );
  }

  return { adminEmail, domain };
}

// ─── Template variables ───────────────────────────────────────────────────────

/**
 * Returns the list of supported template variables that can be embedded in a
 * signature template. Each {{variable}} is replaced with the corresponding
 * value from the user's Directory profile at save/apply time.
 *
 * @returns {Array<{variable:string, label:string, description:string}>}
 */
function getAvailableVariables() {
  return [
    { variable: '{{firstName}}',   label: 'First Name',   description: "User's given name" },
    { variable: '{{lastName}}',    label: 'Last Name',    description: "User's family name" },
    { variable: '{{fullName}}',    label: 'Full Name',    description: "User's full display name" },
    { variable: '{{email}}',       label: 'Email',        description: "User's primary email address" },
    { variable: '{{workPhone}}',   label: 'Work Phone',   description: 'Work phone number from Directory' },
    { variable: '{{mobilePhone}}', label: 'Mobile Phone', description: 'Mobile phone number from Directory' },
    { variable: '{{jobTitle}}',    label: 'Job Title',    description: 'Job title from primary organization' },
    { variable: '{{department}}',  label: 'Department',   description: 'Department from primary organization' },
    { variable: '{{company}}',     label: 'Company',      description: 'Organization name from primary organization' },
  ];
}

// ─── User listing ─────────────────────────────────────────────────────────────

/**
 * Lists up to 100 active users in the domain, ordered by email address.
 * Impersonates ADMIN_EMAIL to call the Admin Directory API.
 *
 * Call repeatedly with the returned nextPageToken to paginate through all users.
 *
 * @param {string|null} [pageToken]  Pagination token from a previous call; null for first page.
 * @returns {{ users: Array<{primaryEmail:string, name:{fullName:string}, suspended:boolean, thumbnailPhotoUrl:string}>, nextPageToken: string|null }}
 */
function getUsers(pageToken) {
  // Uses the AdminDirectory advanced service (runs as the deploying admin's
  // credentials — no service account token needed for this admin-level call).
  var params = {
    customer:   'my_customer',
    maxResults: 100,
    orderBy:    'email',
    // isMailboxSetup=true only for accounts with Gmail provisioned;
    // we surface this in the sidebar so admins can filter non-Gmail accounts.
    fields:
      'users(primaryEmail,name/fullName,thumbnailPhotoUrl,suspended,isMailboxSetup,orgUnitPath),' +
      'nextPageToken'
  };
  if (pageToken) params.pageToken = pageToken;

  var data = AdminDirectory.Users.list(params);
  return {
    users:         data.users         || [],
    nextPageToken: data.nextPageToken || null
  };
}

// ─── Signature retrieval ──────────────────────────────────────────────────────

/**
 * Returns the current HTML signature for a user's primary send-as address.
 * Impersonates the target user via domain-wide delegation.
 *
 * @param {string} userEmail  User's primary email address.
 * @returns {{ signature: string, sendAsEmail: string }}
 */
function getUserSignature(userEmail) {
  const token = getGmailToken_(userEmail);

  const url =
    'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) +
    '/settings/sendAs';

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Gmail API error for ' + userEmail +
      ' (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  const data = JSON.parse(response.getContentText());
  // Prefer the entry flagged as primary; fall back to the first entry
  const primary =
    (data.sendAs || []).find(function(s) { return s.isPrimary; }) ||
    (data.sendAs || [])[0] ||
    {};

  return {
    signature:    primary.signature    || '',
    sendAsEmail:  primary.sendAsEmail  || userEmail
  };
}

// ─── Signature preview ────────────────────────────────────────────────────────

/**
 * Renders the signature template with real Directory data for a specific user,
 * without saving anything. Used by the "Preview" button in the UI.
 *
 * @param {string} userEmail     The user whose profile provides variable values.
 * @param {string} htmlTemplate  Signature template containing optional {{variables}}.
 * @returns {{ html: string, userEmail: string }}
 */
function previewSignatureForUser(userEmail, htmlTemplate) {
  const profile = getUserProfile_(userEmail);
  const vars    = extractVariables_(profile);
  return {
    html:      substituteVariables_(htmlTemplate, vars),
    userEmail: userEmail
  };
}

// ─── Signature updates ────────────────────────────────────────────────────────

/**
 * Fetches the user's Directory profile, substitutes any {{variables}} in the
 * template, then saves the rendered HTML as their Gmail signature.
 *
 * @param {string} userEmail     Target user's primary email address.
 * @param {string} htmlTemplate  HTML template containing optional {{variables}}.
 * @returns {{ success: boolean, email: string }}
 */
function updateUserSignature(userEmail, htmlTemplate) {
  const profile  = getUserProfile_(userEmail);
  const vars     = extractVariables_(profile);
  const rendered = substituteVariables_(htmlTemplate, vars);
  updateSignatureForUser_(userEmail, rendered);
  return { success: true, email: userEmail };
}

/**
 * Updates the Gmail signature for ALL active (non-suspended) users in the domain.
 * For each user, fetches their Directory profile and substitutes {{variables}}
 * before saving, so every user receives a personalised version of the template.
 *
 * Paginates through the full user directory automatically.
 *
 * ⚠ Apps Script execution limit is 6 minutes. For domains with many users
 *   (roughly 500+), this may time out. See README.md → Troubleshooting.
 *
 * @param {string} htmlTemplate  Signature template with optional {{variables}}.
 * @returns {{ updated: number, failed: Array<{email:string, error:string}> }}
 */
function updateAllUsersSignature(htmlTemplate) {
  const results = { updated: 0, failed: [] };
  let pageToken  = null;

  do {
    const page = getUsers(pageToken);
    pageToken = page.nextPageToken;

    for (var i = 0; i < page.users.length; i++) {
      var user = page.users[i];

      // Skip suspended accounts and users without a Gmail mailbox.
      // Attempting to set a signature for a non-Gmail account returns 400
      // "Mail service not enabled" from the Gmail API.
      if (user.suspended || user.isMailboxSetup === false) continue;

      try {
        var profile  = getUserProfile_(user.primaryEmail);
        var vars     = extractVariables_(profile);
        var rendered = substituteVariables_(htmlTemplate, vars);
        updateSignatureForUser_(user.primaryEmail, rendered);
        results.updated++;
      } catch (e) {
        results.failed.push({ email: user.primaryEmail, error: e.message });
      }
    }
  } while (pageToken);

  return results;
}

// ─── Private helpers ──────────────────────────────────────────────────────────

/**
 * Performs the two-step PATCH to update a user's primary send-as signature:
 *   1. GET /sendAs to find the primary sendAsEmail address.
 *   2. PATCH /sendAs/{sendAsEmail} with the new signature.
 *
 * @param {string} userEmail
 * @param {string} htmlSignature
 */
function updateSignatureForUser_(userEmail, htmlSignature) {
  const token = getGmailToken_(userEmail);
  const baseUrl =
    'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) +
    '/settings/sendAs';

  // Step 1 — identify the primary send-as address
  const getResp = UrlFetchApp.fetch(baseUrl, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (getResp.getResponseCode() !== 200) {
    throw new Error(
      'Could not fetch sendAs for ' + userEmail + ': ' +
      getResp.getContentText()
    );
  }

  const sendAsList = JSON.parse(getResp.getContentText()).sendAs || [];
  const primary =
    sendAsList.find(function(s) { return s.isPrimary; }) ||
    sendAsList[0];

  if (!primary) {
    throw new Error('No send-as address found for ' + userEmail);
  }

  // Step 2 — PATCH the signature on the primary address
  const patchUrl = baseUrl + '/' + encodeURIComponent(primary.sendAsEmail);
  const patchResp = UrlFetchApp.fetch(patchUrl, {
    method: 'PATCH',
    headers: {
      Authorization:  'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ signature: htmlSignature }),
    muteHttpExceptions: true
  });

  if (patchResp.getResponseCode() !== 200) {
    throw new Error(
      'Signature update failed for ' + userEmail +
      ' (' + patchResp.getResponseCode() + '): ' +
      patchResp.getContentText()
    );
  }
}

/**
 * Obtains a Gmail-scoped access token that impersonates `userEmail`.
 *
 * @param {string} userEmail
 * @returns {string} OAuth2 access token.
 */
function getGmailToken_(userEmail) {
  return getServiceAccountToken_(userEmail, [SCOPE_GMAIL_SETTINGS_]);
}

/**
 * Fetches a user's full Directory profile, including name parts, phones, and
 * organization info needed for variable substitution.
 *
 * Uses the AdminDirectory advanced service — no service account token needed.
 *
 * @param {string} userEmail  User's primary email address.
 * @returns {object} Raw user resource from the Directory API.
 */
function getUserProfile_(userEmail) {
  return AdminDirectory.Users.get(userEmail, {
    fields: 'primaryEmail,name(givenName,familyName,fullName),phones,organizations'
  });
}

/**
 * Maps a Directory API user resource to a flat { variableName: value } object
 * used for template substitution.
 *
 * Missing Directory fields produce empty strings so that no literal
 * "{{placeholder}}" text appears in the final rendered signature.
 *
 * @param {object} profile  Raw user resource from getUserProfile_().
 * @returns {Object.<string, string>}
 */
function extractVariables_(profile) {
  var name   = profile.name          || {};
  var phones = profile.phones        || [];
  var orgs   = profile.organizations || [];

  // Prefer the entry flagged as primary; fall back to the first entry
  var primaryOrg = orgs.filter(function(o) { return o.primary; })[0] || orgs[0] || {};

  // Returns the value of the first phone matching a given type, or ''
  function phoneByType(type) {
    var match = phones.filter(function(p) { return p.type === type; })[0];
    return match ? (match.value || '') : '';
  }

  return {
    firstName:   name.givenName        || '',
    lastName:    name.familyName       || '',
    fullName:    name.fullName         || '',
    email:       profile.primaryEmail  || '',
    workPhone:   phoneByType('work'),
    mobilePhone: phoneByType('mobile'),
    jobTitle:    primaryOrg.title      || '',
    department:  primaryOrg.department || '',
    company:     primaryOrg.name       || ''
  };
}

/**
 * Replaces every {{variableName}} placeholder in an HTML template string with
 * the corresponding value from the `vars` map.
 *
 * Uses split+join rather than a regex so that special characters in
 * replacement values (e.g. "$" in a phone number) are never misinterpreted.
 *
 * @param {string}               template  HTML template string.
 * @param {Object.<string,string>} vars    Map of variable name → value.
 * @returns {string} Rendered HTML.
 */
function substituteVariables_(template, vars) {
  var result = template;
  Object.keys(vars).forEach(function(key) {
    result = result.split('{{' + key + '}}').join(vars[key]);
  });
  return result;
}

// ─── Data Transfer ────────────────────────────────────────────────────────────

/** Scope for the Admin Data Transfer API. */
const SCOPE_DATATRANSFER_ =
  'https://www.googleapis.com/auth/admin.datatransfer';

// ─── Delegation & Calendar scopes ─────────────────────────────────────────────

/** Scope for managing Gmail delegates (settings.sharing includes read+write). */
const SCOPE_GMAIL_DELEGATION_  =
  'https://www.googleapis.com/auth/gmail.settings.sharing';

/** Scope for managing Google Calendar ACLs on behalf of users. */
const SCOPE_CALENDAR_ =
  'https://www.googleapis.com/auth/calendar';

/** Scope for reading policies via the Cloud Identity Policy API. */
const SCOPE_CLOUD_IDENTITY_POLICIES_ =
  'https://www.googleapis.com/auth/cloud-identity.policies';

/**
 * Returns the list of applications available for data transfer in this domain.
 * The frontend uses the returned app IDs when calling createDataTransfer().
 *
 * Requires the Data Transfer API to be enabled on the GCP project and the
 * SCOPE_DATATRANSFER_ scope to be authorised in Admin Console DWD.
 *
 * @returns {{ applications: Array<{id:string, name:string, transferParams:Array}> }}
 */
function getDataTransferApplications() {
  const { adminEmail } = getConfig();
  const token = getServiceAccountToken_(adminEmail, [SCOPE_DATATRANSFER_]);

  const response = UrlFetchApp.fetch(
    'https://admin.googleapis.com/admin/datatransfer/v1/applications' +
    '?customerId=my_customer',
    {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Data Transfer applications API error (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  const data = JSON.parse(response.getContentText());
  return { applications: data.applications || [] };
}

/**
 * Creates a data transfer from one user to another, covering one or more
 * applications in a single request.
 *
 * @param {string} fromEmail    Source user's primary email address.
 * @param {string} toEmail      Destination user's primary email address.
 * @param {Array}  appTransfers Array of { applicationId: string, params: Array<{key,value}> }
 *                              — one entry per selected service (Drive, Calendar, Looker Studio).
 * @returns {object} The created DataTransfer resource from the API.
 */
function createDataTransfer(fromEmail, toEmail, appTransfers) {
  const { adminEmail } = getConfig();
  const xferToken = getServiceAccountToken_(adminEmail, [SCOPE_DATATRANSFER_]);

  // Data Transfer API requires immutable user IDs, not email addresses.
  // getUserId_ uses AdminDirectory advanced service — no token needed.
  const fromId = getUserId_(fromEmail);
  const toId   = getUserId_(toEmail);

  const payload = {
    oldOwnerUserId: fromId,
    newOwnerUserId: toId,
    applicationDataTransfers: (appTransfers || []).map(function(at) {
      return {
        applicationId:             String(at.applicationId),
        applicationTransferParams: at.params || []
      };
    })
  };

  const response = UrlFetchApp.fetch(
    'https://admin.googleapis.com/admin/datatransfer/v1/transfers',
    {
      method: 'POST',
      headers: {
        Authorization:  'Bearer ' + xferToken,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Transfer creation failed (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
}

/**
 * Returns the current status of an existing data transfer.
 *
 * @param {string} transferId  Transfer ID returned by createDataTransfer().
 * @returns {object} The DataTransfer resource (including current status).
 */
function getDataTransferStatus(transferId) {
  const { adminEmail } = getConfig();
  const token = getServiceAccountToken_(adminEmail, [SCOPE_DATATRANSFER_]);

  const response = UrlFetchApp.fetch(
    'https://admin.googleapis.com/admin/datatransfer/v1/transfers/' +
    encodeURIComponent(transferId),
    {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Could not fetch transfer status (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
}

/**
 * Resolves a user's primary email to their immutable Directory user ID.
 * The Data Transfer API requires the immutable ID rather than an email address.
 *
 * Uses the AdminDirectory advanced service — no service account token needed.
 *
 * @param {string} email  User's primary email address.
 * @returns {string} Immutable user ID.
 */
function getUserId_(email) {
  return AdminDirectory.Users.get(email, { fields: 'id' }).id;
}

// ─── Mail Delegation ──────────────────────────────────────────────────────────

/**
 * Returns the current Gmail delegates for a user.
 * Throws on any API error — the caller is responsible for checking the
 * mail delegation policy (via checkMailDelegationPolicy) when an error occurs.
 *
 * @param {string} userEmail  User whose delegates to list.
 * @returns {{ delegates: Array }}
 */
function getDelegates(userEmail) {
  const token = getServiceAccountToken_(userEmail, [SCOPE_GMAIL_SETTINGS_, SCOPE_GMAIL_DELEGATION_]);

  const response = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) + '/settings/delegates',
    {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  var code = response.getResponseCode();

  // 204 No Content is a valid success — the user simply has no delegates configured.
  if (code === 204) return { delegates: [] };

  if (code !== 200) {
    var rawBody = response.getContentText();
    var reason  = rawBody;
    try {
      var parsed = JSON.parse(rawBody);
      reason = (parsed.error && parsed.error.message) ? parsed.error.message : rawBody;
    } catch (e) { /* keep raw */ }
    throw new Error('Gmail Delegates API error (' + code + '): ' + reason);
  }

  const data = JSON.parse(response.getContentText());
  return { delegates: data.delegates || [] };
}

/**
 * Adds a delegate to a user's Gmail account.
 *
 * @param {string} userEmail     User who is granting delegation access.
 * @param {string} delegateEmail The email address to add as a delegate.
 * @returns {object} The created delegate resource.
 */
function addDelegate(userEmail, delegateEmail) {
  const token = getServiceAccountToken_(userEmail, [SCOPE_GMAIL_SETTINGS_, SCOPE_GMAIL_DELEGATION_]);

  const response = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) + '/settings/delegates',
    {
      method: 'POST',
      headers: {
        Authorization:  'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ delegateEmail: delegateEmail }),
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200 && response.getResponseCode() !== 201) {
    throw new Error(
      'Failed to add delegate (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
}

/**
 * Removes a delegate from a user's Gmail account.
 *
 * @param {string} userEmail     User who owns the mailbox.
 * @param {string} delegateEmail The delegate address to remove.
 * @returns {{ success: boolean }}
 */
function removeDelegate(userEmail, delegateEmail) {
  const token = getServiceAccountToken_(userEmail, [SCOPE_GMAIL_SETTINGS_, SCOPE_GMAIL_DELEGATION_]);

  const response = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) + '/settings/delegates/' +
    encodeURIComponent(delegateEmail),
    {
      method: 'DELETE',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200 && response.getResponseCode() !== 204) {
    throw new Error(
      'Failed to remove delegate (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return { success: true };
}

/**
 * Resolves the mail delegation policy for an organisational unit using the
 * Cloud Identity Policy API (cloudidentity.googleapis.com/v1/policies).
 *
 * Flow:
 *   1. Look up the OU's immutable ID via the Admin Directory API.
 *   2. GET /v1/policies filtered by the mail_delegation setting type and OU.
 *
 * @param {string} orgUnitPath  OU path as returned by the Directory API (e.g. "/Sales").
 * @returns {{ allowed: boolean|null, orgUnitId: string, note: string|null }}
 */
function checkMailDelegationPolicy(orgUnitPath) {
  const { adminEmail } = getConfig();

  // Step 1 — fetch ALL gmail.mail_delegation policies, no OU filter.
  // Filtering by the user's exact OU would miss inherited policies; we handle
  // inheritance ourselves by walking the OU ancestry below (same approach as GAM).
  const policyToken = getServiceAccountToken_(adminEmail, [SCOPE_CLOUD_IDENTITY_POLICIES_]);
  const filter = 'setting.type=="settings/gmail.mail_delegation"';

  const policyResp = UrlFetchApp.fetch(
    'https://cloudidentity.googleapis.com/v1/policies?filter=' + encodeURIComponent(filter),
    {
      headers: { Authorization: 'Bearer ' + policyToken },
      muteHttpExceptions: true
    }
  );

  if (policyResp.getResponseCode() !== 200) {
    return { allowed: null, note: null };
  }

  const allPolicies = JSON.parse(policyResp.getContentText()).policies || [];
  if (!allPolicies.length) return { allowed: null, note: null };

  // Step 2 — resolve each policy's OU resource name ("orgUnits/<id>") to its path
  // using the AdminDirectory advanced service (no service account token needed).
  const ouPathMap = {}; // "orgUnits/<id>" → "/Some/OU/Path"

  for (var i = 0; i < allPolicies.length; i++) {
    var ouResource = allPolicies[i].policyQuery && allPolicies[i].policyQuery.orgUnit;
    if (!ouResource || ouPathMap.hasOwnProperty(ouResource)) continue;

    var ouId = ouResource.replace(/^orgUnits\//, '');
    try {
      var ou = AdminDirectory.Orgunits.get('my_customer', 'id:' + ouId, { fields: 'orgUnitPath' });
      ouPathMap[ouResource] = ou.orgUnitPath || '/';
    } catch (e) {
      ouPathMap[ouResource] = '/';
    }
  }

  // Step 3 — walk the user's OU path from most-specific to root.
  // Return the value of the first matching policy found (most specific wins).
  const userPath = orgUnitPath || '/';
  const segments = userPath.split('/').filter(Boolean);

  for (var depth = segments.length; depth >= 0; depth--) {
    var candidatePath = depth === 0 ? '/' : '/' + segments.slice(0, depth).join('/');

    // Collect all policies whose resolved OU path matches this candidate level.
    var candidates = allPolicies.filter(function(p) {
      return ouPathMap[p.policyQuery && p.policyQuery.orgUnit] === candidatePath;
    });
    if (!candidates.length) continue;

    // Among matching policies, pick the one with the highest sortOrder
    // (ADMIN overrides SYSTEM at the same level; higher sortOrder = more specific).
    var best = candidates.reduce(function(a, b) {
      return ((b.policyQuery.sortOrder || 0) > (a.policyQuery.sortOrder || 0)) ? b : a;
    });

    var value = (best.setting && best.setting.value) || {};
    var allowed = typeof value.enableMailDelegation !== 'undefined'
      ? !!value.enableMailDelegation
      : null;

    return { allowed: allowed, note: null };
  }

  return { allowed: null, note: null };
}

// ─── Calendar ACL ─────────────────────────────────────────────────────────────

/**
 * Checks whether Google Calendar is enabled for a user by attempting to fetch
 * their primary calendar resource.
 *
 * @param {string} userEmail
 * @returns {{ enabled: boolean, summary: string }}
 */
function getCalendarStatus(userEmail) {
  try {
    const token = getServiceAccountToken_(userEmail, [SCOPE_CALENDAR_]);

    const response = UrlFetchApp.fetch(
      'https://www.googleapis.com/calendar/v3/calendars/primary',
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      return { enabled: false, summary: '' };
    }

    const data = JSON.parse(response.getContentText());
    return { enabled: true, summary: data.summary || userEmail };
  } catch (e) {
    return { enabled: false, summary: '' };
  }
}

/**
 * Returns the ACL entries for a user's primary calendar.
 *
 * @param {string} ownerEmail  Calendar owner's email address.
 * @returns {{ items: Array }}
 */
function getCalendarAcl(ownerEmail) {
  const token = getServiceAccountToken_(ownerEmail, [SCOPE_CALENDAR_]);

  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/calendar/v3/calendars/primary/acl',
    {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Calendar ACL API error (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  const data = JSON.parse(response.getContentText());
  return { items: data.items || [] };
}

/**
 * Adds a user-scoped ACL entry to the owner's primary calendar.
 * sendNotifications is always set to false — no email is sent to the grantee.
 *
 * @param {string} ownerEmail    Calendar owner's email address.
 * @param {string} granteeEmail  User to share the calendar with.
 * @param {string} role          One of: freeBusyReader, reader, writer.
 * @returns {object} The created AclRule resource.
 */
function addCalendarAcl(ownerEmail, granteeEmail, role) {
  const token = getServiceAccountToken_(ownerEmail, [SCOPE_CALENDAR_]);

  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/calendar/v3/calendars/primary/acl' +
    '?sendNotifications=false',
    {
      method: 'POST',
      headers: {
        Authorization:  'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        role:  role,
        scope: { type: 'user', value: granteeEmail }
      }),
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200 && response.getResponseCode() !== 201) {
    throw new Error(
      'Failed to add calendar ACL (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
}

/**
 * Removes an ACL rule from the owner's primary calendar.
 *
 * @param {string} ownerEmail  Calendar owner's email address.
 * @param {string} ruleId      ACL rule ID returned by getCalendarAcl().
 * @returns {{ success: boolean }}
 */
function removeCalendarAcl(ownerEmail, ruleId) {
  const token = getServiceAccountToken_(ownerEmail, [SCOPE_CALENDAR_]);

  const response = UrlFetchApp.fetch(
    'https://www.googleapis.com/calendar/v3/calendars/primary/acl/' +
    encodeURIComponent(ruleId),
    {
      method: 'DELETE',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200 && response.getResponseCode() !== 204) {
    throw new Error(
      'Failed to remove calendar ACL (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return { success: true };
}

// ─── License Management ───────────────────────────────────────────────────────
//
// All Licensing API calls use the AdminLicenseManager advanced service
// (enabled in appsscript.json). No service account token needed here —
// the advanced service runs as the deploying admin's credentials.

/**
 * Returns the domain's real customer ID (e.g. "C0abc123") by looking up the
 * deploying admin's account via the Directory API. Falls back to 'my_customer'
 * if the lookup fails. The Licensing API list operations require this ID.
 *
 * @returns {string}
 */
function getCustomerId_() {
  try {
    var admin = AdminDirectory.Users.get(Session.getActiveUser().getEmail());
    return admin.customerId || 'my_customer';
  } catch (e) {
    return 'my_customer';
  }
}

/**
 * SKU IDs that cannot be individually assigned:
 *   – "Archived User" SKUs  (retain data, no active service)
 *   – "Former Employee" vault archive SKU
 *   – Education Fundamentals (site-based blanket license, not per-user)
 */
var NON_ASSIGNABLE_SKU_IDS_ = [
  '1010020021',              // Business Starter – Archived
  '1010020027',              // Business Standard – Archived
  '1010020023',              // Business Plus – Archived
  '1010020029',              // Enterprise Standard – Archived
  '1010020031',              // Enterprise Plus – Archived
  'Google-Vault-Former-Employee', // Vault Former Employee (archive)
  '1010310002',              // Education Fundamentals (site-based)
];

/** Education SKU IDs — excluded from the standard assignment/sync dropdowns. */
var EDUCATION_SKU_IDS_ = [
  '1010310002', // Education Fundamentals (also non-assignable)
  '1010310003', // Education Standard
  '1010310001', // Education Plus – Legacy
  '1010310005', // Education Plus
  '1010310004', // Teaching and Learning Upgrade
];

/**
 * Legacy G Suite SKU IDs and Essentials SKU IDs that cannot coexist with
 * standard Workspace (Business/Enterprise) SKUs — excluded from all dropdowns.
 */
var LEGACY_AND_ESSENTIALS_SKU_IDS_ = [
  'Google-Apps-For-Business', // G Suite Basic (Legacy)
  'Google-Apps-Unlimited',    // G Suite Business (Legacy)
  '1010060001',               // Enterprise Essentials
  '1010060003',               // Essentials
  '1010060005',               // Enterprise Essentials Plus
];

/**
 * Full static catalog of Google Workspace product + SKU combinations.
 * Sourced from:
 * https://developers.google.com/workspace/admin/licensing/v1/how-tos/products
 *
 * @returns {Array<{productId, productName, skus: Array<{skuId, skuName}>}>}
 */
function getLicenseProducts() {
  return [
    {
      productId: 'Google-Apps',
      productName: 'Google Workspace',
      skus: [
        { skuId: '1010020020', skuName: 'Business Starter' },
        { skuId: '1010020021', skuName: 'Business Starter – Archived' },
        { skuId: '1010020025', skuName: 'Business Standard' },
        { skuId: '1010020027', skuName: 'Business Standard – Archived' },
        { skuId: '1010020026', skuName: 'Business Plus' },
        { skuId: '1010020023', skuName: 'Business Plus – Archived' },
        { skuId: '1010020028', skuName: 'Enterprise Standard' },
        { skuId: '1010020029', skuName: 'Enterprise Standard – Archived' },
        { skuId: '1010020030', skuName: 'Enterprise Plus' },
        { skuId: '1010020031', skuName: 'Enterprise Plus – Archived' },
        { skuId: '1010060001', skuName: 'Enterprise Essentials' },
        { skuId: '1010060003', skuName: 'Essentials' },
        { skuId: '1010060005', skuName: 'Enterprise Essentials Plus' },
        { skuId: '1010020034', skuName: 'Frontline Starter' },
        { skuId: '1010020035', skuName: 'Frontline Standard' },
        { skuId: '1010310002', skuName: 'Education Fundamentals' },
        { skuId: '1010310003', skuName: 'Education Standard' },
        { skuId: '1010310001', skuName: 'Education Plus – Legacy' },
        { skuId: '1010310005', skuName: 'Education Plus' },
        { skuId: '1010310004', skuName: 'Teaching and Learning Upgrade' },
        { skuId: 'Google-Apps-For-Business', skuName: 'G Suite Basic (Legacy)' },
        { skuId: 'Google-Apps-Unlimited',    skuName: 'G Suite Business (Legacy)' },
      ]
    },
    {
      productId: 'Google-Vault',
      productName: 'Google Vault',
      skus: [
        { skuId: 'Google-Vault',                skuName: 'Google Vault' },
        { skuId: 'Google-Vault-Former-Employee', skuName: 'Former Employee' },
      ]
    },
    {
      productId: 'Google-Drive-storage',
      productName: 'Google Drive Storage',
      skus: [
        { skuId: 'Google-Drive-storage-20GB',  skuName: '20 GB' },
        { skuId: 'Google-Drive-storage-50GB',  skuName: '50 GB' },
        { skuId: 'Google-Drive-storage-200GB', skuName: '200 GB' },
        { skuId: 'Google-Drive-storage-400GB', skuName: '400 GB' },
        { skuId: 'Google-Drive-storage-1TB',   skuName: '1 TB' },
        { skuId: 'Google-Drive-storage-2TB',   skuName: '2 TB' },
        { skuId: 'Google-Drive-storage-4TB',   skuName: '4 TB' },
        { skuId: 'Google-Drive-storage-8TB',   skuName: '8 TB' },
        { skuId: 'Google-Drive-storage-16TB',  skuName: '16 TB' },
      ]
    },
    {
      productId: 'Cloud-Identity',
      productName: 'Cloud Identity',
      skus: [
        { skuId: '1010050001', skuName: 'Cloud Identity Free' },
        { skuId: '1010050004', skuName: 'Cloud Identity Premium' },
      ]
    },
    {
      productId: 'Google-Chrome-Device-Management',
      productName: 'Chrome Enterprise',
      skus: [
        { skuId: 'Google-Chrome-Device-Management', skuName: 'Chrome Enterprise' },
      ]
    },
  ];
}

/**
 * Returns only product + SKU combinations that have at least one license
 * currently assigned in the domain. Used to populate the Group Sync dropdowns
 * so admins only see what is actually in their environment.
 *
 * Checks each SKU individually via AdminLicenseManager.LicenseAssignments
 * .listForProductAndSku() with maxResults=1 — fast and precise.
 *
 * @returns {Array<{productId, productName, skus: Array<{skuId, skuName}>}>}
 */
function getActiveLicenseProducts() {
  var catalog         = getLicenseProducts();
  var activeProducts  = [];

  catalog.forEach(function(product) {
    var activeSkus = [];

    product.skus.forEach(function(sku) {
      try {
        var result = AdminLicenseManager.LicenseAssignments.listForProductAndSku(
          product.productId, sku.skuId, 'my_customer', { maxResults: 1 }
        );
        if (result.items && result.items.length > 0) {
          activeSkus.push(sku);
        }
      } catch (e) {
        // No assignments for this SKU — skip it.
      }
    });

    if (activeSkus.length > 0) {
      activeProducts.push({
        productId:   product.productId,
        productName: product.productName,
        skus:        activeSkus
      });
    }
  });

  return activeProducts;
}

/**
 * Returns Google Workspace (Google-Apps) SKUs that can be individually
 * assigned, excluding archived SKUs, site-based SKUs, and Education SKUs.
 * Used to populate the per-user assignment picker.
 *
 * @returns {Array<{productId, productName, skus: Array<{skuId, skuName}>}>}
 */
function getAssignableLicenseSkus() {
  var catalog = getLicenseProducts();
  return catalog
    .filter(function(p) { return p.productId === 'Google-Apps'; })
    .map(function(product) {
      return {
        productId:   product.productId,
        productName: product.productName,
        skus: product.skus.filter(function(sku) {
          return NON_ASSIGNABLE_SKU_IDS_.indexOf(sku.skuId) === -1 &&
                 sku.skuName.indexOf('Archived') === -1 &&
                 EDUCATION_SKU_IDS_.indexOf(sku.skuId) === -1 &&
                 LEGACY_AND_ESSENTIALS_SKU_IDS_.indexOf(sku.skuId) === -1;
        })
      };
    })
    .filter(function(p) { return p.skus.length > 0; });
}

/**
 * Returns a flat list of Google Workspace SKUs currently in use in the domain,
 * filtered to only assignable SKUs (no archived, no education, no legacy, no
 * essentials). Builds from getLicenseInventory() so it matches what the
 * Domain License Inventory section shows.
 *
 * @returns {{ skus: Array<{productId, productName, skuId, skuName}> }}
 */
function getActiveWorkspaceSkus() {
  var result = getLicenseInventory();
  var skus = result.inventory.filter(function(item) {
    return NON_ASSIGNABLE_SKU_IDS_.indexOf(item.skuId) === -1 &&
           EDUCATION_SKU_IDS_.indexOf(item.skuId) === -1 &&
           LEGACY_AND_ESSENTIALS_SKU_IDS_.indexOf(item.skuId) === -1 &&
           item.skuName.indexOf('Archived') === -1;
  }).map(function(item) {
    return {
      productId:   item.productId,
      productName: item.productName,
      skuId:       item.skuId,
      skuName:     item.skuName
    };
  });
  return { skus: skus };
}

/**
 * Returns all Google Workspace SKUs with their assigned-user counts, built by
 * paginating listForProduct (one call covers all SKUs) and grouping by skuId.
 * SKU names come from the API response; falls back to the catalog if missing.
 *
 * @returns {{ inventory: Array<{productId, productName, skuId, skuName, count}> }}
 */
function getLicenseInventory() {
  // Build catalog skuId → skuName lookup
  var catalog   = getLicenseProducts();
  var workspace = catalog.filter(function(p) { return p.productId === 'Google-Apps'; })[0];
  var catalogNames = {};
  if (workspace) {
    workspace.skus.forEach(function(sku) { catalogNames[sku.skuId] = sku.skuName; });
  }

  // Paginate listForProduct to count all assignments grouped by skuId
  var customerId = getCustomerId_();
  var skuCounts  = {};
  var skuNames   = {};
  var pageToken  = null;
  do {
    var opts = { maxResults: 1000 };
    if (pageToken) opts.pageToken = pageToken;
    // Let errors propagate — silent breaks were masking failures as empty results
    var data = AdminLicenseManager.LicenseAssignments.listForProduct(
      'Google-Apps', customerId, opts
    );
    (data.items || []).forEach(function(item) {
      skuCounts[item.skuId] = (skuCounts[item.skuId] || 0) + 1;
      if (item.skuName && !skuNames[item.skuId]) skuNames[item.skuId] = item.skuName;
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);

  var inventory = Object.keys(skuCounts).map(function(skuId) {
    return {
      productId:   'Google-Apps',
      productName: 'Google Workspace',
      skuId:       skuId,
      skuName:     catalogNames[skuId] || skuNames[skuId] || skuId,
      count:       skuCounts[skuId]
    };
  });

  inventory.sort(function(a, b) { return b.count - a.count; });
  return { inventory: inventory };
}

/**
 * Returns all license assignments for a user by scanning listForProduct for
 * each known product. This is more reliable than get()-per-SKU because it
 * returns exactly what the API has assigned rather than depending on catalog
 * SKU IDs matching the API's internal IDs.
 *
 * @param {string} userEmail
 * @returns {{ licenses: Array<{productId, productName, skuId, skuName}> }}
 */
function getUserLicenses(userEmail) {
  var emailLower = userEmail.toLowerCase();
  var customerId = getCustomerId_();
  var products   = getLicenseProducts();
  var licenses   = [];

  // Build skuId → catalog name lookup for display
  var catalogNames = {};
  products.forEach(function(p) {
    p.skus.forEach(function(s) { catalogNames[s.skuId] = s.skuName; });
  });

  products.forEach(function(product) {
    var pageToken = null;
    do {
      var opts = { maxResults: 1000 };
      if (pageToken) opts.pageToken = pageToken;
      try {
        // Let errors propagate — silent breaks were masking failures as empty results
        var data = AdminLicenseManager.LicenseAssignments.listForProduct(
          product.productId, customerId, opts
        );
        var found = false;
        (data.items || []).forEach(function(item) {
          if ((item.userId || '').toLowerCase() === emailLower) {
            found = true;
            licenses.push({
              productId:   product.productId,
              productName: product.productName,
              skuId:       item.skuId,
              skuName:     catalogNames[item.skuId] || item.skuName || item.skuId
            });
          }
        });
        // Stop paging this product once the user is found
        pageToken = found ? null : (data.nextPageToken || null);
      } catch (e) {
        var msg = e.message || '';
        if (msg.indexOf('403') !== -1 || msg.indexOf('do not have permission') !== -1) {
          throw new Error(
            'Permission denied: the apps.licensing OAuth scope has not been authorized.'
          );
        }
        // Other errors (invalid productId, not supported) — skip this product only
      }
    } while (pageToken);
  });

  return { licenses: licenses };
}

/**
 * Counts the number of users assigned to a specific product + SKU.
 * Paginates up to 50,000 users; sets capped=true if that limit is reached.
 *
 * @param {string} productId
 * @param {string} skuId
 * @returns {{ count: number, capped: boolean }}
 */
function getLicenseCount(productId, skuId) {
  var count     = 0;
  var capped    = false;
  var pageToken = null;
  var page      = 0;
  var MAX_PAGES = 50;

  do {
    var opts = { maxResults: 1000 };
    if (pageToken) opts.pageToken = pageToken;

    try {
      var result = AdminLicenseManager.LicenseAssignments.listForProductAndSku(
        productId, skuId, 'my_customer', opts
      );
      count    += (result.items || []).length;
      pageToken = result.nextPageToken || null;
      page++;
      if (page >= MAX_PAGES && pageToken) { capped = true; break; }
    } catch (e) {
      break;
    }
  } while (pageToken);

  return { count: count, capped: capped };
}

/**
 * Assigns a license to a user.
 *
 * @param {string} userEmail
 * @param {string} productId
 * @param {string} skuId
 * @returns {object} The created LicenseAssignment resource.
 */
function assignLicense(userEmail, productId, skuId) {
  try {
    return AdminLicenseManager.LicenseAssignments.insert(
      { userId: userEmail }, productId, skuId
    );
  } catch (e) {
    if (e.message && e.message.indexOf('412') !== -1) {
      throw new Error(
        'License cannot be assigned — the user\'s OU may have automatic licensing ' +
        'enabled. Disable auto-licensing for their OU in the Admin Console first. ' +
        'Details: ' + e.message
      );
    }
    throw new Error('Assign license failed: ' + e.message);
  }
}

/**
 * Switches a user from one SKU to another within the same product.
 *
 * @param {string} userEmail
 * @param {string} productId   Must be the same product for both SKUs.
 * @param {string} oldSkuId    Current SKU.
 * @param {string} newSkuId    Target SKU.
 * @returns {object} The updated LicenseAssignment resource.
 */
function switchLicense(userEmail, productId, oldSkuId, newSkuId) {
  try {
    return AdminLicenseManager.LicenseAssignments.update(
      { skuId: newSkuId }, productId, oldSkuId, userEmail
    );
  } catch (e) {
    if (e.message && e.message.indexOf('412') !== -1) {
      throw new Error(
        'License cannot be switched — the user\'s OU may have automatic licensing ' +
        'enabled. Details: ' + e.message
      );
    }
    throw new Error('Switch license failed: ' + e.message);
  }
}

/**
 * Removes a license from a user.
 *
 * @param {string} userEmail
 * @param {string} productId
 * @param {string} skuId
 * @returns {{ success: boolean }}
 */
function unassignLicense(userEmail, productId, skuId) {
  try {
    AdminLicenseManager.LicenseAssignments.remove(productId, skuId, userEmail);
    return { success: true };
  } catch (e) {
    throw new Error('Unassign license failed: ' + e.message);
  }
}

/**
 * Lists all groups in the domain (up to 1,000).
 *
 * @returns {{ groups: Array<{email: string, name: string}> }}
 */
function getGroups() {
  var groups    = [];
  var pageToken = null;

  do {
    var params = {
      customer:   'my_customer',
      maxResults: 200,
      orderBy:    'email',
      fields:     'groups(email,name),nextPageToken'
    };
    if (pageToken) params.pageToken = pageToken;

    var data  = AdminDirectory.Groups.list(params);
    groups    = groups.concat(data.groups || []);
    pageToken = data.nextPageToken || null;
  } while (pageToken && groups.length < 1000);

  return { groups: groups };
}

/**
 * Returns the email addresses of all USER members of a group.
 * includeDerivedMembership=true expands nested groups transitively.
 *
 * @param {string} groupEmail
 * @returns {{ members: string[] }}
 */
function getGroupMembers(groupEmail) {
  var members   = [];
  var pageToken = null;

  do {
    var params = {
      maxResults:               200,
      includeDerivedMembership: true,
      fields:                   'members(email,type,status),nextPageToken'
    };
    if (pageToken) params.pageToken = pageToken;

    var data = AdminDirectory.Members.list(groupEmail, params);
    (data.members || []).forEach(function(m) {
      if (m.type === 'USER' && m.status === 'ACTIVE') {
        members.push(m.email.toLowerCase());
      }
    });
    pageToken = data.nextPageToken || null;
  } while (pageToken);

  return { members: members };
}

/**
 * Runs or previews a license sync for a group.
 *
 * Sync rules:
 *   1. User IN group, has target SKU              → NO_CHANGE
 *   2. User IN group, has different SKU (product) → SWITCH
 *   3. User IN group, no license in product       → ASSIGN
 *   4. User NOT in group, has license in product  → REMOVE
 *
 * @param {string}  groupEmail  Group whose members define the licensed set.
 * @param {string}  productId   Product family to manage.
 * @param {string}  skuId       Target SKU to enforce.
 * @param {boolean} dryRun      If true, returns proposed changes without applying them.
 * @returns {{ changes: Array, applied: boolean }}
 */
function syncLicensesForGroup(groupEmail, productId, skuId, dryRun) {
  // Resolve group membership
  var memberResult = getGroupMembers(groupEmail);
  var memberSet    = {};
  memberResult.members.forEach(function(email) { memberSet[email] = true; });

  // Get all current licensees for this product (any SKU)
  var currentLicensees = getLicensedUsersForProduct_(productId);

  // Build change list
  var changes = [];

  // Process group members
  Object.keys(memberSet).forEach(function(email) {
    var currentSku = currentLicensees[email];
    if (!currentSku) {
      changes.push({ action: 'ASSIGN',    email: email, fromSku: null,       toSku: skuId });
    } else if (currentSku === skuId) {
      changes.push({ action: 'NO_CHANGE', email: email, fromSku: currentSku, toSku: skuId });
    } else {
      changes.push({ action: 'SWITCH',    email: email, fromSku: currentSku, toSku: skuId });
    }
  });

  // Process licensed users NOT in the group
  Object.keys(currentLicensees).forEach(function(email) {
    if (!memberSet[email]) {
      changes.push({ action: 'REMOVE', email: email, fromSku: currentLicensees[email], toSku: null });
    }
  });

  if (dryRun) {
    return { changes: changes, applied: false };
  }

  // Apply changes
  var results = changes.map(function(change) {
    if (change.action === 'NO_CHANGE') {
      return mergeObj_(change, { status: 'skipped' });
    }
    try {
      if (change.action === 'ASSIGN') {
        assignLicense(change.email, productId, skuId);
      } else if (change.action === 'SWITCH') {
        switchLicense(change.email, productId, change.fromSku, skuId);
      } else if (change.action === 'REMOVE') {
        unassignLicense(change.email, productId, change.fromSku);
      }
      return mergeObj_(change, { status: 'success' });
    } catch (e) {
      return mergeObj_(change, { status: 'error', error: e.message });
    }
  });

  return { changes: results, applied: true };
}

/**
 * Writes a sync change list to a Google Sheet.
 * Creates a new spreadsheet if sheetId is null; appends to existing otherwise.
 *
 * @param {string}      groupEmail
 * @param {string}      productId
 * @param {string}      skuId
 * @param {Array}       changes     Change array from syncLicensesForGroup().
 * @param {string|null} sheetId     Existing spreadsheet ID, or null to create new.
 * @returns {{ spreadsheetId: string, spreadsheetUrl: string }}
 */
function writeSyncLog(groupEmail, productId, skuId, changes, sheetId) {
  // Uses SpreadsheetApp (built-in service) — no service account token needed.
  var sheetTab = 'License Sync';
  var ss;

  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    // Create a new spreadsheet in the deploying admin's Drive.
    ss = SpreadsheetApp.create(
      'License Sync Log \u2013 ' + new Date().toISOString().slice(0, 10)
    );
    var defaultSheet = ss.getSheets()[0];
    defaultSheet.setName(sheetTab);
    defaultSheet.appendRow([
      'Timestamp', 'Group', 'Product ID', 'SKU ID',
      'User Email', 'Action', 'From SKU', 'To SKU', 'Status', 'Error'
    ]);
  }

  var sheet = ss.getSheetByName(sheetTab);
  if (!sheet) {
    // Tab doesn't exist yet in an existing spreadsheet — create it with header.
    sheet = ss.insertSheet(sheetTab);
    sheet.appendRow([
      'Timestamp', 'Group', 'Product ID', 'SKU ID',
      'User Email', 'Action', 'From SKU', 'To SKU', 'Status', 'Error'
    ]);
  }

  var ts = new Date().toISOString();
  changes
    .filter(function(c) { return c.action !== 'NO_CHANGE'; })
    .forEach(function(c) {
      sheet.appendRow([
        ts, groupEmail, productId, skuId,
        c.email, c.action,
        c.fromSku || '', c.toSku || '',
        c.status  || 'dry_run', c.error || ''
      ]);
    });

  return {
    spreadsheetId: ss.getId(),
    url:           ss.getUrl()
  };
}

// ── Private helpers ───────────────────────────────────────────────────────────

/**
 * Returns { lowerEmail: skuId } for all users licensed under productId (any SKU).
 * Uses AdminLicenseManager advanced service.
 *
 * @param {string} productId
 * @returns {Object}
 */
function getLicensedUsersForProduct_(productId) {
  var result    = {};
  var pageToken = null;

  do {
    var opts = { maxResults: 1000 };
    if (pageToken) opts.pageToken = pageToken;

    try {
      var data = AdminLicenseManager.LicenseAssignments.listForProduct(
        productId, 'my_customer', opts
      );
      (data.items || []).forEach(function(item) {
        result[item.userId.toLowerCase()] = item.skuId;
      });
      pageToken = data.nextPageToken || null;
    } catch (e) {
      break;
    }
  } while (pageToken);

  return result;
}

/**
 * Shallow-merges two plain objects (V8-compatible replacement for Object.assign).
 *
 * @param {Object} base
 * @param {Object} extra
 * @returns {Object}
 */
function mergeObj_(base, extra) {
  var result = {};
  Object.keys(base).forEach(function(k)  { result[k] = base[k];  });
  Object.keys(extra).forEach(function(k) { result[k] = extra[k]; });
  return result;
}
