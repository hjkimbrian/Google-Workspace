/**
 * Tests.gs — GWS Admin Toolkit test suite
 *
 * Run from the Apps Script editor:
 *   Functions dropdown → runAllTests → Run
 *
 * Tests are split into two groups:
 *   Unit tests     — pure functions, no network calls, always runnable.
 *   Integration tests — call live APIs; require valid Script Properties and
 *                       AdminDirectory / Licensing / etc. to be enabled.
 *
 * Each test function throws on failure and returns normally on success.
 * runAllTests() catches exceptions and logs a PASS/FAIL summary.
 */

// ════════════════════════════════════════════════════════════════════════════
//  Assertion helpers
// ════════════════════════════════════════════════════════════════════════════

function assert_(condition, message) {
  if (!condition) {
    throw new Error('Assertion failed' + (message ? ': ' + message : ''));
  }
}

function assertEqual_(actual, expected, message) {
  if (actual !== expected) {
    throw new Error(
      (message || 'assertEqual') +
      ' — expected ' + JSON.stringify(expected) +
      ', got '      + JSON.stringify(actual)
    );
  }
}

function assertDeepEqual_(actual, expected, message) {
  var a = JSON.stringify(actual);
  var b = JSON.stringify(expected);
  if (a !== b) {
    throw new Error(
      (message || 'assertDeepEqual') + ' — expected ' + b + ', got ' + a
    );
  }
}

function assertThrows_(fn, msgContains, label) {
  var threw = false;
  try { fn(); } catch (e) {
    threw = true;
    if (msgContains && e.message.indexOf(msgContains) === -1) {
      throw new Error(
        (label || 'assertThrows') +
        ': exception message "' + e.message +
        '" does not contain "' + msgContains + '"'
      );
    }
  }
  if (!threw) throw new Error((label || 'assertThrows') + ': expected an exception but none was thrown');
}

// ════════════════════════════════════════════════════════════════════════════
//  Test runner
// ════════════════════════════════════════════════════════════════════════════

/**
 * Runs every function whose name starts with "test_" in this file,
 * collecting PASS/FAIL results and logging a summary.
 *
 * To run: open Apps Script editor → select runAllTests → click Run.
 */
function runAllTests() {
  var allTests = [
    // ── Unit tests (no API calls) ────────────────────────────────────────
    test_mergeObj_basic,
    test_mergeObj_overwrite,
    test_mergeObj_empty,
    test_substituteVariables_single,
    test_substituteVariables_multiple,
    test_substituteVariables_missing_key,
    test_substituteVariables_no_placeholders,
    test_extractVariables_full_profile,
    test_extractVariables_empty_profile,
    test_extractVariables_primary_org,
    test_getLicenseProducts_structure,
    test_getLicenseProducts_skus_nonempty,
    // ── Integration tests (require valid config + API access) ────────────
    test_getConfig_has_required_fields,
    test_getUsers_returns_array,
    test_getUsers_pagination_token,
    test_getUsers_fields_present,
    test_getUserProfile_returns_name,
    test_getUserId_returns_string,
    test_getGroups_returns_array,
    test_getGroupMembers_invalid_group_throws,
    test_getLicenseCount_returns_number,
    test_syncLicensesForGroup_dryrun_structure
  ];

  var passed = 0;
  var failed = 0;
  var log    = [];

  allTests.forEach(function(fn) {
    try {
      fn();
      log.push('PASS  ' + fn.name);
      passed++;
    } catch (e) {
      log.push('FAIL  ' + fn.name + '  —  ' + e.message);
      failed++;
    }
  });

  log.push('');
  log.push('Results: ' + passed + ' passed, ' + failed + ' failed out of ' + allTests.length);

  log.forEach(function(line) { Logger.log(line); });

  if (failed > 0) {
    throw new Error(failed + ' test(s) failed — see Logger output for details.');
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  Unit tests — mergeObj_
// ════════════════════════════════════════════════════════════════════════════

function test_mergeObj_basic() {
  var result = mergeObj_({ a: 1 }, { b: 2 });
  assertEqual_(result.a, 1, 'a preserved');
  assertEqual_(result.b, 2, 'b added');
}

function test_mergeObj_overwrite() {
  var result = mergeObj_({ a: 1, b: 2 }, { b: 99 });
  assertEqual_(result.b, 99, 'b overwritten');
  assertEqual_(result.a, 1,  'a preserved');
}

function test_mergeObj_empty() {
  assertDeepEqual_(mergeObj_({}, {}), {}, 'two empty objects');
  assertEqual_(mergeObj_({ x: 5 }, {}).x, 5, 'empty extra');
}

// ════════════════════════════════════════════════════════════════════════════
//  Unit tests — substituteVariables_
// ════════════════════════════════════════════════════════════════════════════

function test_substituteVariables_single() {
  var result = substituteVariables_('Hello {{firstName}}!', { firstName: 'Ada' });
  assertEqual_(result, 'Hello Ada!');
}

function test_substituteVariables_multiple() {
  var result = substituteVariables_(
    '{{firstName}} {{lastName}} <{{email}}>',
    { firstName: 'Ada', lastName: 'Lovelace', email: 'ada@example.com' }
  );
  assertEqual_(result, 'Ada Lovelace <ada@example.com>');
}

function test_substituteVariables_missing_key() {
  // A placeholder with no matching key should remain unchanged so admins can
  // spot typos rather than silently getting an empty field.
  var result = substituteVariables_('Hello {{firstName}}!', {});
  assertEqual_(result, 'Hello {{firstName}}!', 'unknown key left intact');
}

function test_substituteVariables_no_placeholders() {
  var result = substituteVariables_('<b>Static content</b>', { firstName: 'Ada' });
  assertEqual_(result, '<b>Static content</b>', 'no-op when no placeholders');
}

// ════════════════════════════════════════════════════════════════════════════
//  Unit tests — extractVariables_
// ════════════════════════════════════════════════════════════════════════════

function test_extractVariables_full_profile() {
  var profile = {
    primaryEmail: 'ada@example.com',
    name:  { givenName: 'Ada', familyName: 'Lovelace', fullName: 'Ada Lovelace' },
    phones: [
      { type: 'work',   value: '+1-555-0100' },
      { type: 'mobile', value: '+1-555-0199' }
    ],
    organizations: [
      { primary: true, title: 'Engineer', department: 'R&D', name: 'Acme Corp' }
    ]
  };
  var vars = extractVariables_(profile);
  assertEqual_(vars.firstName,   'Ada',          'firstName');
  assertEqual_(vars.lastName,    'Lovelace',     'lastName');
  assertEqual_(vars.fullName,    'Ada Lovelace', 'fullName');
  assertEqual_(vars.email,       'ada@example.com', 'email');
  assertEqual_(vars.workPhone,   '+1-555-0100',  'workPhone');
  assertEqual_(vars.mobilePhone, '+1-555-0199',  'mobilePhone');
  assertEqual_(vars.jobTitle,    'Engineer',     'jobTitle');
  assertEqual_(vars.department,  'R&D',          'department');
  assertEqual_(vars.company,     'Acme Corp',    'company');
}

function test_extractVariables_empty_profile() {
  var vars = extractVariables_({});
  // Every variable should be an empty string — no {{placeholder}} leakage.
  var keys = ['firstName','lastName','fullName','email','workPhone','mobilePhone','jobTitle','department','company'];
  keys.forEach(function(k) {
    assertEqual_(vars[k], '', k + ' is empty string for empty profile');
  });
}

function test_extractVariables_primary_org() {
  // When multiple org entries exist, the primary one should win.
  var profile = {
    organizations: [
      { primary: false, title: 'Old Title', department: 'Old', name: 'OldCo' },
      { primary: true,  title: 'CEO',       department: 'Exec', name: 'NewCo' }
    ]
  };
  var vars = extractVariables_(profile);
  assertEqual_(vars.jobTitle,   'CEO',   'primary org title');
  assertEqual_(vars.company,    'NewCo', 'primary org company');
  assertEqual_(vars.department, 'Exec',  'primary org department');
}

// ════════════════════════════════════════════════════════════════════════════
//  Unit tests — getLicenseProducts (static catalog, no API call)
// ════════════════════════════════════════════════════════════════════════════

function test_getLicenseProducts_structure() {
  var products = getLicenseProducts();
  assert_(Array.isArray(products), 'returns an array');
  assert_(products.length > 0, 'catalog is non-empty');
  products.forEach(function(p) {
    assert_(typeof p.productId   === 'string' && p.productId,   'productId is a non-empty string');
    assert_(typeof p.productName === 'string' && p.productName, 'productName is a non-empty string');
    assert_(Array.isArray(p.skus) && p.skus.length > 0,         'skus is non-empty array');
    p.skus.forEach(function(s) {
      assert_(typeof s.skuId   === 'string' && s.skuId,   'skuId is a non-empty string');
      assert_(typeof s.skuName === 'string' && s.skuName, 'skuName is a non-empty string');
    });
  });
}

function test_getLicenseProducts_skus_nonempty() {
  // Every product must have at least one SKU so the UI dropdown is never empty.
  var products = getLicenseProducts();
  products.forEach(function(p) {
    assert_(p.skus.length >= 1, p.productName + ' has at least one SKU');
  });
}

// ════════════════════════════════════════════════════════════════════════════
//  Integration tests
//  These call live Google APIs. They require:
//    • Valid ADMIN_EMAIL and DOMAIN Script Properties (getConfig works)
//    • AdminDirectory advanced service enabled in the Apps Script project
//    • Service account configured with the DWD scopes listed in README.md
// ════════════════════════════════════════════════════════════════════════════

function test_getConfig_has_required_fields() {
  var cfg = getConfig();
  assert_(typeof cfg.adminEmail === 'string' && cfg.adminEmail, 'adminEmail present');
  assert_(typeof cfg.domain     === 'string' && cfg.domain,     'domain present');
  assert_(cfg.adminEmail.indexOf('@') > 0, 'adminEmail looks like an email address');
}

function test_getUsers_returns_array() {
  var result = getUsers(null);
  assert_(Array.isArray(result.users), 'users is an array');
  assert_(result.users.length > 0, 'at least one user returned');
}

function test_getUsers_pagination_token() {
  var first = getUsers(null);
  // nextPageToken should be a string (more pages) or null (last page).
  assert_(
    first.nextPageToken === null || typeof first.nextPageToken === 'string',
    'nextPageToken is null or string'
  );
}

function test_getUsers_fields_present() {
  var result = getUsers(null);
  var user   = result.users[0];
  assert_(typeof user.primaryEmail === 'string', 'primaryEmail present');
  assert_(user.name && typeof user.name.fullName === 'string', 'name.fullName present');
  assert_(typeof user.suspended === 'boolean', 'suspended is boolean');
}

function test_getUserProfile_returns_name() {
  // Uses the first user from the directory to avoid hardcoding an email.
  var firstUser = getUsers(null).users[0];
  var profile   = getUserProfile_(firstUser.primaryEmail);
  assert_(profile.name && profile.name.fullName, 'name.fullName populated');
  assertEqual_(profile.primaryEmail, firstUser.primaryEmail, 'email matches');
}

function test_getUserId_returns_string() {
  var firstUser = getUsers(null).users[0];
  var id        = getUserId_(firstUser.primaryEmail);
  assert_(typeof id === 'string' && id.length > 0, 'id is a non-empty string');
}

function test_getGroups_returns_array() {
  var result = getGroups();
  assert_(Array.isArray(result.groups), 'groups is an array');
  // A domain may have zero groups — just verify the shape.
  result.groups.forEach(function(g) {
    assert_(typeof g.email === 'string' && g.email, 'group email present');
  });
}

function test_getGroupMembers_invalid_group_throws() {
  // An email that is definitely not a valid group should cause AdminDirectory
  // to throw — verify we propagate the error rather than silently returning [].
  assertThrows_(
    function() { getGroupMembers('this-group-does-not-exist-xyz@invalid.example'); },
    null,
    'test_getGroupMembers_invalid_group_throws'
  );
}

function test_getLicenseCount_returns_number() {
  // Use the first product/SKU from the static catalog.
  var products = getLicenseProducts();
  var product  = products[0];
  var sku      = product.skus[0];
  var result   = getLicenseCount(product.productId, sku.skuId);
  assert_(typeof result.count === 'number', 'count is a number');
  assert_(result.count >= 0, 'count is non-negative');
}

function test_syncLicensesForGroup_dryrun_structure() {
  // Pick the first group (if any); if no groups exist, skip gracefully.
  var groupsResult = getGroups();
  if (groupsResult.groups.length === 0) {
    Logger.log('test_syncLicensesForGroup_dryrun_structure: SKIP (no groups in domain)');
    return;
  }

  var group    = groupsResult.groups[0];
  var products = getLicenseProducts();
  var product  = products[0];
  var sku      = product.skus[0];

  var result = syncLicensesForGroup(group.email, product.productId, sku.skuId, true);

  assert_(Array.isArray(result.changes), 'changes is an array');
  assertEqual_(result.applied, false, 'dryRun: applied is false');

  var validActions = ['ASSIGN', 'SWITCH', 'REMOVE', 'NO_CHANGE'];
  result.changes.forEach(function(c) {
    assert_(validActions.indexOf(c.action) !== -1, 'action is a valid value: ' + c.action);
    assert_(typeof c.email === 'string' && c.email, 'change has email: ' + JSON.stringify(c));
  });
}
