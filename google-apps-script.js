/**
 * Google Apps Script Webhook Handler for Speak Up Day ICPR 2026
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName('Pendaftar') || doc.getActiveSheet();

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 16).getValues()[0];
    if (headers.length === 0 || headers[0] === '') {
      headers = [
        'Timestamp',
        'Ticket ID',
        'Nama Lengkap',
        'NISN',
        'Sekolah',
        'Kelas',
        'WhatsApp',
        'Email',
        'Cabang Lomba',
        'Anggota 2',
        'Anggota 3',
        'Guru Pendamping',
        'Instagram',
        'Link Kartu Pelajar',
        'Link Bukti Bayar',
        'Status'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    var data = JSON.parse(e.postData.contents);
    var newRow = [
      new Date(),
      data.ticketId || '',
      data.fullName || '',
      data.nisn || '',
      data.school || '',
      data.grade || '',
      data.whatsapp || '',
      data.email || '',
      data.arenaTitle || '',
      data.member2 || '',
      data.member3 || '',
      data.teacher || '',
      data.instagram || '',
      data.proofCard || '',
      data.proofPayment || '',
      data.status || 'Menunggu Verifikasi'
    ];

    sheet.appendRow(newRow);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
