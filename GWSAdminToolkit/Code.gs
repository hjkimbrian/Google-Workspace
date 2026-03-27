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

/** Scope for listing users via the Admin Directory API. */
const SCOPE_DIRECTORY_READ_ =
  'https://www.googleapis.com/auth/admin.directory.user.readonly';

/** Scope for reading and writing Gmail send-as settings (signatures). */
const SCOPE_GMAIL_SETTINGS_ =
  'https://www.googleapis.com/auth/gmail.settings.basic';

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
  const { adminEmail, domain } = getConfig();
  const token = getServiceAccountToken_(adminEmail, [SCOPE_DIRECTORY_READ_]);

  // Build query string manually for V8/Apps Script compatibility.
  //
  // Use customer=my_customer instead of domain= so that users from ALL domains
  // in the Workspace account are returned (primary domain + any secondary /
  // alias domains such as strataprimedemo.com alongside demo.strataprime.com).
  // The domain= parameter restricts results to a single domain only.
  let url =
    'https://admin.googleapis.com/admin/directory/v1/users' +
    '?customer=my_customer' +
    '&maxResults=100' +
    '&orderBy=email' +
    // Request only the fields we need to minimise response size
    '&fields=' + encodeURIComponent(
      // isMailboxSetup = true only for accounts that have Gmail provisioned;
      // we surface this in the sidebar so admins can filter non-Gmail accounts.
      'users(primaryEmail,name/fullName,thumbnailPhotoUrl,suspended,isMailboxSetup,orgUnitPath),' +
      'nextPageToken'
    );

  if (pageToken) {
    url += '&pageToken=' + encodeURIComponent(pageToken);
  }

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Directory API error (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  const data = JSON.parse(response.getContentText());
  return {
    users: data.users || [],
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
  const { adminEmail } = getConfig();

  // Obtain one directory token to reuse across all profile fetches;
  // valid for 1 hour (sufficient for domains up to ~500 users).
  const dirToken = getServiceAccountToken_(adminEmail, [SCOPE_DIRECTORY_READ_]);

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
        var profile  = getUserProfile_(user.primaryEmail, dirToken);
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
 * @param {string} userEmail   User's primary email address.
 * @param {string} [dirToken]  Pre-fetched directory token. When calling in a
 *                             loop, pass a cached token to avoid one token
 *                             request per user. Omit for one-off calls.
 * @returns {object} Raw user resource from the Directory API.
 */
function getUserProfile_(userEmail, dirToken) {
  var token = dirToken;
  if (!token) {
    var adminEmail = getConfig().adminEmail;
    token = getServiceAccountToken_(adminEmail, [SCOPE_DIRECTORY_READ_]);
  }

  var url =
    'https://admin.googleapis.com/admin/directory/v1/users/' +
    encodeURIComponent(userEmail) +
    '?fields=' + encodeURIComponent(
      'primaryEmail,name(givenName,familyName,fullName),phones,organizations'
    );

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Could not fetch profile for ' + userEmail +
      ' (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText());
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

/** Scope for reading organisational unit records from the Directory API. */
const SCOPE_DIRECTORY_ORGUNIT_ =
  'https://www.googleapis.com/auth/admin.directory.orgunit.readonly';

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

  // Obtain both tokens up front; each JWT exchange counts against quota.
  const dirToken  = getServiceAccountToken_(adminEmail, [SCOPE_DIRECTORY_READ_]);
  const xferToken = getServiceAccountToken_(adminEmail, [SCOPE_DATATRANSFER_]);

  // Data Transfer API requires immutable user IDs, not email addresses.
  const fromId = getUserId_(fromEmail, dirToken);
  const toId   = getUserId_(toEmail,   dirToken);

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
 * @param {string} email     User's primary email address.
 * @param {string} dirToken  Pre-fetched Directory API token (optional).
 * @returns {string} Immutable user ID.
 */
function getUserId_(email, dirToken) {
  var token = dirToken;
  if (!token) {
    token = getServiceAccountToken_(getConfig().adminEmail, [SCOPE_DIRECTORY_READ_]);
  }

  var response = UrlFetchApp.fetch(
    'https://admin.googleapis.com/admin/directory/v1/users/' +
    encodeURIComponent(email) + '?fields=id',
    {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'User not found: ' + email +
      ' (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText()).id;
}

// ─── Mail Delegation ──────────────────────────────────────────────────────────

/**
 * Returns the current Gmail delegates for a user and the delegation policy
 * status for their organisational unit.
 *
 * @param {string}      userEmail    User whose delegates to list.
 * @param {string|null} orgUnitPath  User's OU path (from Directory API).
 *                                   Pass null to skip the policy check.
 * @returns {{ delegates: Array, delegationAllowed: boolean|null, policyNote: string|null }}
 */
function getDelegates(userEmail, orgUnitPath) {
  const token = getServiceAccountToken_(userEmail, [SCOPE_GMAIL_DELEGATION_]);

  const response = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/' +
    encodeURIComponent(userEmail) + '/settings/delegates',
    {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (response.getResponseCode() === 403) {
    throw new Error(
      'Permission denied reading Gmail delegates for ' + userEmail + '. ' +
      'Add the scope https://www.googleapis.com/auth/gmail.settings.sharing ' +
      'to your service account\'s Domain-Wide Delegation entry in ' +
      'Admin Console → Security → API Controls → Domain-wide Delegation, ' +
      'then retry.'
    );
  }

  if (response.getResponseCode() !== 200) {
    throw new Error(
      'Gmail Delegates API error (' + response.getResponseCode() + '): ' +
      response.getContentText()
    );
  }

  const data = JSON.parse(response.getContentText());

  var policyResult = { allowed: null, note: null };
  if (orgUnitPath) {
    try {
      policyResult = checkMailDelegationPolicy(orgUnitPath);
    } catch (e) {
      policyResult = { allowed: null, note: 'Could not check policy: ' + e.message };
    }
  }

  return {
    delegates:         data.delegates || [],
    delegationAllowed: policyResult.allowed,
    policyNote:        policyResult.note
  };
}

/**
 * Adds a delegate to a user's Gmail account.
 *
 * @param {string} userEmail     User who is granting delegation access.
 * @param {string} delegateEmail The email address to add as a delegate.
 * @returns {object} The created delegate resource.
 */
function addDelegate(userEmail, delegateEmail) {
  const token = getServiceAccountToken_(userEmail, [SCOPE_GMAIL_DELEGATION_]);

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
  const token = getServiceAccountToken_(userEmail, [SCOPE_GMAIL_DELEGATION_]);

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

  // Step 1 — resolve OU path → OU ID via Directory API
  const dirToken = getServiceAccountToken_(adminEmail, [SCOPE_DIRECTORY_ORGUNIT_]);

  const cleanPath = orgUnitPath.replace(/^\//, ''); // strip leading slash
  const ouResp = UrlFetchApp.fetch(
    'https://admin.googleapis.com/admin/directory/v1/customer/my_customer/orgunits/' +
    encodeURIComponent(cleanPath),
    {
      headers: { Authorization: 'Bearer ' + dirToken },
      muteHttpExceptions: true
    }
  );

  if (ouResp.getResponseCode() !== 200) {
    return {
      allowed: null,
      orgUnitId: '',
      note: 'Could not resolve OU: ' + ouResp.getContentText()
    };
  }

  var orgUnitId = (JSON.parse(ouResp.getContentText()).orgUnitId || '').replace(/^id:/, '');
  if (!orgUnitId) {
    return { allowed: null, orgUnitId: '', note: 'OU ID not found.' };
  }

  // Step 2 — query Cloud Identity Policy API for mail delegation setting on this OU.
  // Filter matches any setting type containing "mail_delegation" for resilience against
  // namespace changes (e.g. gmail.mail_delegation vs workspace.gmail.mail_delegation).
  const policyToken = getServiceAccountToken_(adminEmail, [SCOPE_CLOUD_IDENTITY_POLICIES_]);

  const filter =
    'setting.type.matches(\'.*mail_delegation\')' +
    ' && policyQuery.orgUnit=="orgUnits/' + orgUnitId + '"';

  const policyResp = UrlFetchApp.fetch(
    'https://cloudidentity.googleapis.com/v1/policies' +
    '?filter=' + encodeURIComponent(filter),
    {
      headers: { Authorization: 'Bearer ' + policyToken },
      muteHttpExceptions: true
    }
  );

  if (policyResp.getResponseCode() !== 200) {
    return {
      allowed: null,
      orgUnitId: orgUnitId,
      note: 'Policy API error: ' + policyResp.getContentText()
    };
  }

  const policyData = JSON.parse(policyResp.getContentText());
  const policies   = policyData.policies || [];

  if (policies.length === 0) {
    return {
      allowed: null,
      orgUnitId: orgUnitId,
      note: 'No explicit policy set (inheriting from parent OU).'
    };
  }

  // Parse the setting value — try common field names for the delegation boolean.
  var allowed = null;
  for (var i = 0; i < policies.length; i++) {
    var val = policies[i].setting && policies[i].setting.value;
    if (!val) continue;

    if (typeof val.mailDelegationEnabled !== 'undefined') {
      allowed = !!val.mailDelegationEnabled;
      break;
    }
    if (typeof val.enableMailDelegation !== 'undefined') {
      allowed = !!val.enableMailDelegation;
      break;
    }
    // Fallback: first boolean field in the value object
    var keys = Object.keys(val);
    for (var j = 0; j < keys.length; j++) {
      if (typeof val[keys[j]] === 'boolean') {
        allowed = val[keys[j]];
        break;
      }
    }
    if (allowed !== null) break;
  }

  return { allowed: allowed, orgUnitId: orgUnitId, note: null };
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
