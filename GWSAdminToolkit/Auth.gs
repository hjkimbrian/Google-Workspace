/**
 * SignatureManager — Auth.gs
 *
 * Service account authentication via domain-wide delegation (DWD).
 *
 * This module creates a signed RS256 JWT from the service account credentials
 * stored in Script Properties, then exchanges it for a short-lived OAuth2
 * access token. The private key is NEVER hardcoded — it lives in
 * Script Properties only.
 *
 * Required Script Property:
 *   SERVICE_ACCOUNT_KEY — Full JSON content of the downloaded service account
 *                         key file (the entire contents of the .json file).
 */

const TOKEN_ENDPOINT_ = 'https://oauth2.googleapis.com/token';

/**
 * Returns an OAuth2 access token that impersonates `subjectEmail`, using the
 * service account key stored in Script Properties.
 *
 * @param {string}   subjectEmail  The Workspace user to impersonate.
 * @param {string[]} scopes        Array of OAuth2 scope URLs to request.
 * @returns {string} A valid access token (expires in 1 hour).
 */
function getServiceAccountToken_(subjectEmail, scopes) {
  const keyJson = PropertiesService.getScriptProperties()
    .getProperty('SERVICE_ACCOUNT_KEY');

  if (!keyJson) {
    throw new Error(
      'SERVICE_ACCOUNT_KEY not set in Script Properties. ' +
      'See README.md → Step 4 for setup instructions.'
    );
  }

  let credentials;
  try {
    credentials = JSON.parse(keyJson);
  } catch (e) {
    throw new Error(
      'SERVICE_ACCOUNT_KEY is not valid JSON. ' +
      'Paste the entire contents of the service account .json file.'
    );
  }

  const jwt = buildJwt_(credentials, subjectEmail, scopes);
  return exchangeJwtForToken_(jwt);
}

// ─── Private helpers ──────────────────────────────────────────────────────────

/**
 * Builds a signed RS256 JWT for a service account impersonation request.
 *
 * JWT spec: https://tools.ietf.org/html/rfc7519
 * Google service account auth: https://developers.google.com/identity/protocols/oauth2/service-account
 *
 * @param {object}   credentials   Parsed service account JSON key object.
 * @param {string}   subjectEmail  User to impersonate (the "sub" claim).
 * @param {string[]} scopes        OAuth2 scopes to request.
 * @returns {string} Compact serialized JWT: header.claims.signature
 */
function buildJwt_(credentials, subjectEmail, scopes) {
  const now = Math.floor(Date.now() / 1000);

  // ── Header ────────────────────────────────────────────────────────────────
  const header = base64UrlEncode_(JSON.stringify({
    alg: 'RS256',
    typ: 'JWT'
  }));

  // ── Claim set ─────────────────────────────────────────────────────────────
  const claimSet = base64UrlEncode_(JSON.stringify({
    iss: credentials.client_email, // Service account identity
    sub: subjectEmail,             // Workspace user to impersonate (DWD)
    scope: scopes.join(' '),       // Requested OAuth2 scopes
    aud: TOKEN_ENDPOINT_,          // Token endpoint as audience
    iat: now,                      // Issued at (Unix epoch seconds)
    exp: now + 3600                // Expires in 1 hour (max allowed)
  }));

  const signingInput = `${header}.${claimSet}`;

  // ── Signature ─────────────────────────────────────────────────────────────
  // Utilities.computeRsaSha256Signature accepts the PEM private_key string
  // directly and returns a Byte[].
  const signatureBytes = Utilities.computeRsaSha256Signature(
    signingInput,
    credentials.private_key
  );

  return `${signingInput}.${base64UrlEncode_(signatureBytes)}`;
}

/**
 * POSTs a signed JWT to Google's token endpoint and returns the access token.
 *
 * @param {string} jwt  Signed compact JWT.
 * @returns {string} access_token string.
 */
function exchangeJwtForToken_(jwt) {
  const response = UrlFetchApp.fetch(TOKEN_ENDPOINT_, {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      `Token exchange failed (HTTP ${response.getResponseCode()}): ` +
      response.getContentText()
    );
  }

  return JSON.parse(response.getContentText()).access_token;
}

/**
 * Encodes a UTF-8 string or Byte[] to Base64url (no padding, URL-safe chars)
 * as required by the JWT spec (RFC 7515 §2).
 *
 * @param {string|Byte[]} data
 * @returns {string} Base64url-encoded string.
 */
function base64UrlEncode_(data) {
  const encoded = (typeof data === 'string')
    ? Utilities.base64Encode(data, Utilities.Charset.UTF_8)
    : Utilities.base64Encode(data);

  // Standard base64 → base64url: replace +→-, /→_, strip trailing =
  return encoded.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}
