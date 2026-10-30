function doGet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ambil pesan
  const pesanSheet = ss.getSheetByName("pesan");
  const pesan = pesanSheet.getRange("A2").getValue();

  // Ambil jadwal
  const jadwalSheet = ss.getSheetByName("jadwal");

  const data = jadwalSheet.getDataRange().getValues();

  const headers = data.shift();

  const jadwal = data.map(row => ({
    tanggal: Utilities.formatDate(
        new Date(row[0]),
        "Asia/Jakarta",
        "dd/MM/yyyy"
    ),
    waktu: row[1],
    tempat: row[2]
  }));

  // Response JSON
  const result = {
    pesan: pesan,
    jadwal: jadwal
  };

  return ContentService
    .createTextOutput(
      JSON.stringify(result)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}