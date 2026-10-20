// Google Sheet > Extensions > Apps Script > dán file này > Deploy > Web app (Execute as: Me, Access: Anyone).
// Dán URL /exec vào LOG_URL trong index.html. Sheet: cột A thời gian, B tên, C link ảnh.
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const clean = v => /^[=+\-@]/.test(v) ? "'" + v : v;       // chặn formula injection
  const name = clean(String(d.name || '').slice(0, 24));
  let url = '';
  if (d.photo && d.photo.length < 2e7) {
    const blob = Utilities.newBlob(Utilities.base64Decode(d.photo.split(',')[1]), 'image/jpeg', Date.now() + '.jpg');
    url = DriveApp.createFile(blob).getUrl();
  }
  SpreadsheetApp.getActiveSheet().appendRow([new Date(), name, url]);
  return ContentService.createTextOutput('ok');
}
