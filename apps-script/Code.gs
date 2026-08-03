/**
 * Lendmate Pro — Job Application backend.
 *
 * Setup:
 * 1. Create a Google Sheet (any name).
 * 2. Extensions > Apps Script, delete the default code, paste this whole file in.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the Web App URL and paste it into CONFIG.APPS_SCRIPT_URL in index.html.
 *
 * On first submission this will auto-create an "Applications" tab with headers,
 * and a "Job Application Resumes" folder in your Drive for uploaded resumes.
 */

const RESUME_FOLDER_NAME = 'Job Application Resumes';
const SHEET_NAME = 'Applications';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    let resumeUrl = '';
    if (data.resumeBase64 && data.resumeFileName) {
      resumeUrl = saveResume(data.resumeBase64, data.resumeFileName, data.resumeMimeType, data.fullName);
    }

    sheet.appendRow([
      new Date(),
      data.fullName || '',
      data.email || '',
      data.phone || '',
      data.position || '',
      data.linkedinProfile || '',
      data.portfolio || '',
      data.howHeard || '',
      data.coverLetter || '',
      data.followedLinkedIn ? 'Yes' : 'No',
      resumeUrl,
      data.availability || '',
      data.relevantExperience || ''
    ]);

    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.message });
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_NAME);
  sheet.appendRow([
    'Timestamp', 'Full Name', 'Email', 'Phone', 'Position',
    'LinkedIn Profile', 'Portfolio / Website', 'How They Heard About Us',
    'Why They Want To Work Here', 'Followed LinkedIn Page', 'Resume Link',
    'Can Commit To Hours', 'Relevant Experience / Links'
  ]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 13);
  return sheet;
}

function saveResume(base64Data, fileName, mimeType, applicantName) {
  const folders = DriveApp.getFoldersByName(RESUME_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(RESUME_FOLDER_NAME);

  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const safeName = (applicantName || 'applicant').replace(/[^a-zA-Z0-9]+/g, '_');
  blob.setName(safeName + ' - ' + fileName);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
