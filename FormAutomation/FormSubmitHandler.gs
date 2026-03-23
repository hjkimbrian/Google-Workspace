/**
 * FormSubmitHandler.gs
 *
 * Triggered on Google Form submission.
 * 1. Creates a Shared Drive named after the request.
 * 2. Creates three Google Groups (managers, contributors, viewers)
 *    with the Shared Drive ID embedded in the group email addresses.
 * 3. Adds the form respondent as an OWNER of each group.
 *
 * Setup:
 *  - Open the bound Apps Script project (or a standalone project linked to the form).
 *  - In the script editor go to Triggers > Add Trigger and configure:
 *      Function: onFormSubmit
 *      Event source: From form
 *      Event type: On form submit
 *  - Enable the following Advanced Services (Resources > Advanced Google Services):
 *      Admin SDK Directory API  (AdminDirectory)
 *      Drive API v3             (Drive)
 *  - The executing account must be a Google Workspace super-admin (or have the
 *    delegated privileges to create Shared Drives and Google Groups).
 *
 * Form field expectations (case-insensitive title matching):
 *  - "Project Name"  – used to name the Shared Drive and groups.
 *  - "Your Email"    – the requestor's Google Workspace email address.
 *    (Alternatively, set READ_EMAIL_FROM_SESSION = true to pull the email
 *     from the form respondent's session automatically.)
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

/** Set to true to read the respondent's email from their login session instead
 *  of a form field.  Requires the form to collect sign-in information.       */
var READ_EMAIL_FROM_SESSION = false;

/** Domain used to build group email addresses, e.g. "example.com".
 *  Update this before deploying.                                              */
var DOMAIN = "example.com";

/** Column titles in the form response sheet (case-insensitive).              */
var FIELD_PROJECT_NAME = "project name";
var FIELD_REQUESTOR_EMAIL = "your email";

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

/**
 * Main trigger function. Attach this to the form's "On form submit" event.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 */
function onFormSubmit(e) {
  var respondentEmail = getRespondentEmail(e);
  var projectName     = getFieldValue(e, FIELD_PROJECT_NAME);

  if (!respondentEmail) {
    throw new Error("Could not determine the requestor's email address.");
  }
  if (!projectName) {
    throw new Error('Form response is missing the "' + FIELD_PROJECT_NAME + '" field.');
  }

  // 1. Create Shared Drive ---------------------------------------------------
  var driveId = createSharedDrive(projectName);
  Logger.log("Created Shared Drive: %s  (ID: %s)", projectName, driveId);

  // 2. Create the three Google Groups ----------------------------------------
  var groups = createProjectGroups(projectName, driveId);
  Logger.log("Created groups: %s", JSON.stringify(groups));

  // 3. Add requestor as manager (OWNER role) to each group -------------------
  addOwnerToGroups(respondentEmail, groups);
  Logger.log("Added %s as OWNER to all groups.", respondentEmail);

  // Optional: record results back to the spreadsheet -------------------------
  recordResults(e, driveId, groups);
}

// ---------------------------------------------------------------------------
// Step 1 – Shared Drive
// ---------------------------------------------------------------------------

/**
 * Creates a Shared Drive with the given display name.
 * Uses Drive Advanced Service (v3).
 * @param {string} name  Display name for the Shared Drive.
 * @returns {string}     The newly created Shared Drive ID.
 */
function createSharedDrive(name) {
  // A unique request ID prevents duplicate drives on retries.
  var requestId = Utilities.getUuid();

  var drive = Drive.Drives.insert({ name: name }, requestId);
  return drive.id;
}

// ---------------------------------------------------------------------------
// Step 2 – Google Groups
// ---------------------------------------------------------------------------

/**
 * Creates managers, contributors, and viewers groups for the project.
 * @param {string} projectName  Human-readable project name.
 * @param {string} driveId      Shared Drive ID (embedded in email addresses).
 * @returns {{ managers: string, contributors: string, viewers: string }}
 *          Object mapping role names to group email addresses.
 */
function createProjectGroups(projectName, driveId) {
  var slug = slugify(projectName);

  var groupDefs = [
    { role: "managers",     email: slug + "-managers-"     + driveId + "@" + DOMAIN,
      name: projectName + " – Managers" },
    { role: "contributors", email: slug + "-contributors-" + driveId + "@" + DOMAIN,
      name: projectName + " – Contributors" },
    { role: "viewers",      email: slug + "-viewers-"      + driveId + "@" + DOMAIN,
      name: projectName + " – Viewers" },
  ];

  var result = {};

  groupDefs.forEach(function(def) {
    var group = AdminDirectory.Groups.insert({
      email: def.email,
      name:  def.name,
    });
    result[def.role] = group.email;
    Logger.log("Created group [%s]: %s", def.role, group.email);
  });

  return result;
}

// ---------------------------------------------------------------------------
// Step 3 – Add requestor as group manager
// ---------------------------------------------------------------------------

/**
 * Adds the requestor as an OWNER member of each group.
 * OWNER is the highest member role and grants full group-management rights.
 * @param {string} email   Requestor's Google Workspace email.
 * @param {{ managers: string, contributors: string, viewers: string }} groups
 */
function addOwnerToGroups(email, groups) {
  Object.keys(groups).forEach(function(role) {
    var groupEmail = groups[role];
    AdminDirectory.Members.insert(
      {
        email: email,
        role:  "OWNER",
      },
      groupEmail
    );
    Logger.log("Added %s as OWNER to %s", email, groupEmail);
  });
}

// ---------------------------------------------------------------------------
// Result recording (optional)
// ---------------------------------------------------------------------------

/**
 * Appends Drive ID and group emails to extra columns in the response sheet.
 * No-ops gracefully if the trigger event is not spreadsheet-backed.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 * @param {string} driveId
 * @param {{ managers: string, contributors: string, viewers: string }} groups
 */
function recordResults(e, driveId, groups) {
  try {
    var range = e.range;
    if (!range) return;

    var sheet = range.getSheet();
    var row   = range.getRow();

    // Write to the columns immediately after the last form column.
    var lastCol = sheet.getLastColumn();
    var headers = [
      "Shared Drive ID",
      "Managers Group",
      "Contributors Group",
      "Viewers Group",
    ];
    var values = [
      driveId,
      groups.managers,
      groups.contributors,
      groups.viewers,
    ];

    // Add header row if this is the first response.
    if (row === 2) {
      headers.forEach(function(h, i) {
        sheet.getRange(1, lastCol + 1 + i).setValue(h);
      });
    }

    values.forEach(function(v, i) {
      sheet.getRange(row, lastCol + 1 + i).setValue(v);
    });
  } catch (err) {
    Logger.log("recordResults (non-fatal): %s", err.message);
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Extracts the respondent's email from the form event or their session.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 * @returns {string|null}
 */
function getRespondentEmail(e) {
  if (READ_EMAIL_FROM_SESSION) {
    return Session.getActiveUser().getEmail() || null;
  }
  return getFieldValue(e, FIELD_REQUESTOR_EMAIL);
}

/**
 * Case-insensitive lookup of a form field value by its item title.
 * Works whether the trigger is bound to a Form or a Spreadsheet.
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 * @param {string} fieldTitle
 * @returns {string|null}
 */
function getFieldValue(e, fieldTitle) {
  var needle = fieldTitle.toLowerCase();

  // Spreadsheet-backed form: e.namedValues is a plain object.
  if (e.namedValues) {
    var keys = Object.keys(e.namedValues);
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].toLowerCase() === needle) {
        var val = e.namedValues[keys[i]];
        return Array.isArray(val) ? val[0] : val;
      }
    }
  }

  // Form-native trigger: e.response is a FormResponse object.
  if (e.response) {
    var items = e.response.getItemResponses();
    for (var j = 0; j < items.length; j++) {
      if (items[j].getItem().getTitle().toLowerCase() === needle) {
        return items[j].getResponse();
      }
    }
  }

  return null;
}

/**
 * Converts a project name to a URL/email-safe slug.
 * @param {string} name
 * @returns {string}
 */
function slugify(name) {
  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .substring(0, 40); // keep group emails under the 64-char local-part limit
}
