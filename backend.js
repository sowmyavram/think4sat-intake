/**
 * Think4SAT Student Intake — Apps Script Web App backend.
 *
 * SETUP:
 * 1. Open your Google Sheet.
 * 2. Extensions -> Apps Script.
 * 3. Delete any existing code and paste this in.
 * 4. Replace RECAPTCHA_SECRET_KEY below with your own secret key from
 *    https://www.google.com/recaptcha/admin (create a "v2 Checkbox" key —
 *    the site key goes in the HTML file, the secret key goes here).
 *    The key currently in this file is Google's public TEST secret key,
 *    which always passes verification. Fine for testing the flow end to
 *    end; provides zero real bot protection until you swap in your own.
 * 5. Save (Ctrl+S / Cmd+S).
 * 6. Click "Deploy" -> "Manage deployments" -> pencil/edit icon on your
 *    existing deployment (or "New deployment" if you don't have one yet).
 *      - Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *      - Version: "New version" (important every time you change this code)
 * 7. Click Deploy. Copy the Web App URL (ends in /exec).
 * 8. Paste that URL into SCRIPT_URL in the HTML form if it changed.
 * 9. Visit the /exec URL directly in an incognito window. You should see
 *    "Think4SAT intake endpoint is running." with no Google login prompt.
 *    If you see a login prompt or "Sorry, unable to open the file" instead,
 *    access is NOT public yet — go back to step 6.
 */

// TEST KEY — replace with your own secret key from the reCAPTCHA admin
// console before going live. Must match the site key used in the HTML file.
const RECAPTCHA_SECRET_KEY = "6LeIxAcTAAAAAGG-vFI1TnRWxMZNFuojJ4WifJWe";

// Exact field names the form sends, in the order you want them as columns.
// This MUST match the `name="..."` attributes in the HTML form's inputs.
const FIELDS = [
  "student_name", "preferred_name", "school", "grade_level",
  "parent_name", "contact", "program_track", "gpa",
  "current_courses", "strong_subjects", "weak_subjects",
  "testing_plan", "test_scores", "competitions",
  "free_time", "deep_interests", "self_directed_project",
  "interest_area", "influences", "activities_list", "leadership",
  "awards", "work_research", "time_commitment",
  "signature_activity", "languages", "first_gen",
  "family_context", "geo_preference", "budget",
  "scale_teamwork", "scale_detail", "scale_risk",
  "handles_challenge", "three_words", "possible_majors",
  "career_interest", "dream_schools", "success_definition",
  "help_areas", "additional_notes"
];

// Minimum time between submissions carrying the SAME contact value (email or
// phone). This is a light server-side brake against rapid repeat/DDoS-style
// posting straight to this URL (bypassing the HTML page and its client-side
// checks entirely). It does not block different people submitting normally.
const MIN_SECONDS_BETWEEN_SAME_CONTACT = 30;

function doPost(e) {
  try {
    const params = e.parameter || {};

    // --- reCAPTCHA verification ---------------------------------------
    // Confirms the request actually came from a human who solved the
    // widget on the real form — not a script POSTing directly to this URL.
    const captchaToken = params['g-recaptcha-response'];
    if (!captchaToken) {
      return jsonOutput({ result: "rejected", reason: "missing captcha" });
    }
    const verifyResponse = UrlFetchApp.fetch(
      'https://www.google.com/recaptcha/api/siteverify',
      {
        method: 'post',
        payload: { secret: RECAPTCHA_SECRET_KEY, response: captchaToken }
      }
    );
    const verifyResult = JSON.parse(verifyResponse.getContentText());
    if (!verifyResult.success) {
      return jsonOutput({ result: "rejected", reason: "captcha failed" });
    }

    // --- Honeypot check ---------------------------------------------------
    // The HTML form has a hidden "website" field real users never see or
    // fill. Anything posted directly to this URL that doesn't set it to
    // empty is either a bot filling every field blindly, or a request that
    // didn't come from the real form at all.
    if (params.website && String(params.website).trim() !== "") {
      return jsonOutput({ result: "rejected", reason: "honeypot" });
    }

    // --- Required fields ---------------------------------------------------
    // Reject anything missing the two fields the real form always requires.
    // Blocks empty/garbage POSTs sent straight at this URL.
    if (!params.student_name || !params.contact) {
      return jsonOutput({ result: "rejected", reason: "missing required fields" });
    }

    // --- Basic rate limiting by contact value ------------------------------
    // Uses Apps Script's CacheService (shared, temporary key/value store) to
    // block the same contact value from writing more than once within the
    // cooldown window, regardless of which browser/device it comes from.
    const cache = CacheService.getScriptCache();
    const cacheKey = 'submit_' + String(params.contact).trim().toLowerCase();
    if (cache.get(cacheKey)) {
      return jsonOutput({ result: "rejected", reason: "rate limited" });
    }
    cache.put(cacheKey, "1", MIN_SECONDS_BETWEEN_SAME_CONTACT);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Make sure the header row exists and matches FIELDS. If the sheet is
    // empty, write the header row automatically on first submission.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp"].concat(FIELDS));
    }

    const row = [new Date()];

    FIELDS.forEach(function (field) {
      let value = params[field];
      if (value === undefined || value === null || value === "") {
        value = "";
      }
      // Checkbox groups (like help_areas) can arrive as a single string or,
      // if the client ever sends repeated keys, Apps Script only keeps the
      // last one in e.parameter — e.parameters (plural) has all of them.
      if (e.parameters && e.parameters[field] && e.parameters[field].length > 1) {
        value = e.parameters[field].join(", ");
      }
      row.push(value);
    });

    sheet.appendRow(row);

    return jsonOutput({ result: "success" });

  } catch (err) {
    return jsonOutput({ result: "error", message: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Lets you sanity-check the deployment by just visiting the URL in a browser.
function doGet(e) {
  return ContentService
    .createTextOutput("Think4SAT intake endpoint is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}
