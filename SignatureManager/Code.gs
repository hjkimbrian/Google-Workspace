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
    .setTitle('Gmail Signature Manager')
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

  // Build query string manually for V8/Apps Script compatibility
  let url =
    'https://admin.googleapis.com/admin/directory/v1/users' +
    '?domain='     + encodeURIComponent(domain) +
    '&maxResults=100' +
    '&orderBy=email' +
    // Request only the fields we need to minimise response size
    '&fields=' + encodeURIComponent(
      'users(primaryEmail,name/fullName,thumbnailPhotoUrl,suspended),' +
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

      // Skip suspended accounts — they cannot receive mail and the API would error
      if (user.suspended) continue;

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
