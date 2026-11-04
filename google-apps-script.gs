/**
 * PRAYAS Enquiry Form — Google Apps Script backend
 *
 * What this does:
 *  1. Receives the form data your website posts.
 *  2. Appends it as a new row in this Google Sheet.
 *  3. Emails a notification to whoever you list in NOTIFY_EMAILS.
 *
 * SETUP:
 *  1. Create a new Google Sheet. In row 1, add these headers exactly:
 *     Timestamp | Name | Email | Phone | Organisation | Role | Message
 *  2. In that Sheet: Extensions -> Apps Script.
 *  3. Delete any starter code, paste this whole file in, and click Save.
 *  4. Update NOTIFY_EMAILS below with your real email address(es).
 *  5. Click Deploy -> New deployment -> gear icon -> "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Click Deploy. Google will ask you to authorize — allow it
 *     (you'll see an "unsafe" warning because it's your own unpublished
 *     script; click Advanced -> Go to [project name] (unsafe) -> Allow).
 *  6. Copy the "Web app URL" you're given.
 *  7. Paste that URL into index.html where it says
 *     PASTE_YOUR_GOOGLE_APPS_SCRIPT_URL_HERE (search for that text).
 *  8. Re-upload index.html. Test the form once end-to-end.
 *
 * NOTE: Every time you click "Deploy -> New deployment" after editing
 * this script, you get a NEW url. If you only need to tweak the code,
 * use "Deploy -> Manage deployments -> Edit -> New version" instead,
 * so the URL your website already uses keeps working.
 */

var NOTIFY_EMAILS = "admissions@prayas.org"; // comma-separate multiple: "a@x.com,b@y.com"

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.email || "",
      data.phone || "",
      data.organisation || "",
      data.role || "",
      data.message || ""
    ]);

    MailApp.sendEmail({
      to: NOTIFY_EMAILS,
      subject: "New PRAYAS Enquiry — " + (data.role || "Unspecified") + " — " + (data.name || "Unknown"),
      body:
        "A new enquiry was submitted on the PRAYAS website:\n\n" +
        "Name: " + (data.name || "-") + "\n" +
        "Email: " + (data.email || "-") + "\n" +
        "Phone: " + (data.phone || "-") + "\n" +
        "Organisation: " + (data.organisation || "-") + "\n" +
        "Role: " + (data.role || "-") + "\n" +
        "Message: " + (data.message || "-") + "\n\n" +
        "Full log: " + SpreadsheetApp.getActiveSpreadsheet().getUrl()
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
