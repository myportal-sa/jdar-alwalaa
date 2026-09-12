/**
 * كود الربط بين الصفحة وجوجل شيت.
 * الصف الأول في الشيت يجب أن يكون: الاسم | الرسالة | التاريخ
 */

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // تجاهل صف العناوين

  const entries = rows
    .filter(r => r[0] || r[1])
    .map(r => ({
      name: String(r[0] || ''),
      message: String(r[1] || ''),
      date: r[2] ? new Date(r[2]).toISOString() : ''
    }))
    .reverse(); // الأحدث أولاً

  return ContentService
    .createTextOutput(JSON.stringify(entries))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const data = JSON.parse(e.postData.contents);

  const name = (data.name || '').toString().slice(0, 60);
  const message = (data.message || '').toString().slice(0, 300);

  sheet.appendRow([name, message, new Date()]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
