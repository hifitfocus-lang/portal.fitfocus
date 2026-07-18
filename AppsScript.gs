/**
 * FitFocus Dashboard — Apps Script backend
 * =========================================
 * This is the ENTIRE backend. No Vercel/Node server needed — Google hosts this for free,
 * and it runs under your own Google account, so it can read your private Excel file in
 * Drive without that file ever being shared publicly.
 *
 * SETUP (one-time):
 * 1. Go to https://script.google.com -> New project. Delete the default code, paste this
 *    whole file in.
 * 2. setup() below is already filled in with your email + password. Just edit
 *    DRIVE_FILE_ID — open your file in Drive, copy the long string between /d/ and /view
 *    in the URL bar, paste it in.
 * 3. In the function dropdown at the top of the editor, select "setup", then click ▶ Run.
 *    Approve the permissions prompt (it's asking to read your own Drive — normal).
 *    Check the execution log to confirm "Done." — this stores your password's HASH (not the
 *    plaintext password) and your file ID in Script Properties.
 * 4. Deploy -> New deployment -> gear icon -> Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    ("Anyone" here just means anyone can reach the URL and get a login prompt — they still
 *    can't get any data without your email + password. This is required because a real
 *    per-user Google-account restriction would force *you* to also authenticate via Google
 *    Sign-In on the frontend, which is the more complex option we didn't pick.)
 * 5. Copy the deployment URL (ends in /exec). Paste it into APPS_SCRIPT_URL near the top of
 *    FitFocusDashboard_final.jsx.
 * 6. To change email/password later: edit the values in setup() and re-run it. You do NOT
 *    need to redeploy for that — Script Properties update immediately.
 */

function setup() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty("DASHBOARD_EMAIL", "management@fitfocus.com");
  props.setProperty("DASHBOARD_PASSWORD_HASH", sha256("Azzazz123"));
  props.setProperty("DRIVE_FILE_ID", "PASTE_YOUR_DRIVE_FILE_ID_HERE");
  Logger.log("Done. Email + password hash + file ID saved to Script Properties.");
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ ok: false, error: "bad request" });
  }

  var props = PropertiesService.getScriptProperties();
  var action = body.action;

  if (action === "login") {
    var storedEmail = props.getProperty("DASHBOARD_EMAIL");
    var storedHash = props.getProperty("DASHBOARD_PASSWORD_HASH");
    if (!storedEmail || !storedHash) return jsonOut({ ok: false, error: "server not configured — run setup() first" });

    var submittedEmail = (body.email || "").trim().toLowerCase();
    var submittedHash = sha256(body.password || "");
    if (submittedEmail !== storedEmail.toLowerCase() || submittedHash !== storedHash) {
      return jsonOut({ ok: false, error: "wrong email or password" });
    }

    var token = Utilities.getUuid();
    // 21600 seconds (6 hours) is the maximum CacheService allows — this is the session length.
    CacheService.getScriptCache().put(token, "valid", 21600);
    return jsonOut({ ok: true, token: token });
  }

  if (action === "getData") {
    var token = body.token;
    var cached = token ? CacheService.getScriptCache().get(token) : null;
    if (cached !== "valid") return jsonOut({ ok: false, error: "session expired, please log in again" });

    var fileId = props.getProperty("DRIVE_FILE_ID");
    if (!fileId) return jsonOut({ ok: false, error: "DRIVE_FILE_ID not configured — run setup() first" });

    try {
      var file = DriveApp.getFileById(fileId);
      var bytes = file.getBlob().getBytes();
      var base64 = Utilities.base64Encode(bytes);
      return jsonOut({ ok: true, fileBase64: base64, fileName: file.getName() });
    } catch (err) {
      return jsonOut({ ok: false, error: "could not read file: " + err.message });
    }
  }

  return jsonOut({ ok: false, error: "unknown action" });
}

function sha256(str) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return raw
    .map(function (b) {
      var v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? "0" + v : v;
    })
    .join("");
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
