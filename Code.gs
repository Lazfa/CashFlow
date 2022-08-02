function onFormSubmit(event) {
  var record_array = [];

  var form = FormApp.openById('aaabbbccc'); // Ganti dengan ID Google Form pada status editing
  var formResponses = form.getResponses();
  var formCount = formResponses.length;

  var formResponse = formResponses[formCount - 1];
  var itemResponses = formResponse.getItemResponses();

  for (var i = 0; i < itemResponses.length; i++) {
      var itemResponse = itemResponses[i];
      var title = itemResponse.getItem().getTitle();
      var answer = itemResponse.getResponse();

      Logger.log(title);
      Logger.log(answer);

      record_array.push(answer);
    }

// Tanggal	Arus	Bentuk	Besaran	Keterangan	Jenis laporan	Laporan dari	Email	Timestamp
  if (itemResponses[1].getResponse() == "Masuk"){
    addRecord(record_array[0], record_array[1], record_array[2], record_array[3], record_array[4], null, null, null);
  } else if (itemResponses[1].getResponse() == "Keluar"){
    addRecord(record_array[0], record_array[1], record_array[2], "-"+record_array[3], record_array[4], null, null, null);
  } else if (itemResponses[1].getResponse() == "Mutasi"){
    addRecord(record_array[0], "Keluar", record_array[2], "-"+record_array[4], "MUTASI_"+record_array[5], null, null, null);
    addRecord(record_array[0], "Masuk", record_array[3], record_array[4], "MUTASI_"+record_array[5], null, null, null);
  } else if (itemResponses[1].getResponse() == "Minta report"){
    sendEmail(record_array[0], record_array[2], record_array[3], record_array[4]);
    addRecord(record_array[0], "Report",null, null, null, record_array[2], record_array[3], record_array[4]);
  } else if (itemResponses[1].getResponse() == "Reset bulanan"){
    var sisa = ambilSisaBulanan();
    addRecord(record_array[0], "Keluar", "Bulanan", "-"+sisa, "BULANAN_"+record_array[3], null, null, null);
    addRecord(record_array[0], "Masuk", "Bank 1", sisa, "BULANAN_"+record_array[3], null, null, null);
    addRecord(record_array[0], "Masuk", "Bulanan", record_array[2], "BULANAN_"+record_array[3], null, null, null);
  }
}

function addRecord(tgl, arus, bentuk, besaran, keterangan, jenis_laporan, laporan_dari, email) {
  var url = 'aaabbbccc';   //ganti dengan URL dari GOOGLE SHEET;
  var ss= SpreadsheetApp.openByUrl(url);
  var dataSheet = ss.getSheetByName("Sheet1");
  dataSheet.appendRow([tgl, arus, bentuk, besaran, keterangan, jenis_laporan, laporan_dari, email, new Date()]);

  Logger.log('Data berhasil ditambahkan');
}

function sendEmail(tgl, jenis_laporan, laporan_dari, email){
  var url = 'aaabbbccc';   //ganti dengan URL dari GOOGLE SHEET;
  var ss= SpreadsheetApp.openByUrl(url);
  var ds = ss.getSheetByName("Sheet2");
  var subjek = "Laporan Keuangan";

  if (jenis_laporan == "Satu Sumber"){
    var isi = 0;

    // Cash, BNI hak, BNI bel, BRI, Gopay, OVO, Bulanan
    if(laporan_dari == "Cash"){
      isi = ds.getRange("B2").getValue();
    } else if (laporan_dari == "Bank 1"){
      isi = ds.getRange("B3").getValue();
    } else if (laporan_dari == "Bank 2"){
      isi = ds.getRange("B4").getValue();
    } else if (laporan_dari == "Bank 3"){
      isi = ds.getRange("B5").getValue();
    } else if (laporan_dari == "Bank 4"){
      isi = ds.getRange("B6").getValue();
    } else if (laporan_dari == "Bank 5") {
      isi = ds.getRange("B7").getValue();
    } else if (laporan_dari == "Bank 6"){
      isi = ds.getRange("B8").getValue();
    }

    var body = "Sumber keuangan: " + laporan_dari + "\rJumlah saldo: " + isi + "\r\rTanggal Request: " + tgl + "\rTimestamp: " + new Date();

  } else if (jenis_laporan == "Semua Sumber"){

    var body = "Sumber Keuangan ---> Jumlah Saldo\r\r" +
      ds.getRange("A2").getValue() + "   --->   " + ds.getRange("B2").getValue() + "\r" +
      ds.getRange("A3").getValue() + "   --->   " + ds.getRange("B3").getValue() + "\r" +
      ds.getRange("A4").getValue() + "   --->   " + ds.getRange("B4").getValue() + "\r" +
      ds.getRange("A5").getValue() + "   --->   " + ds.getRange("B5").getValue() + "\r" +
      ds.getRange("A6").getValue() + "   --->   " + ds.getRange("B6").getValue() + "\r" +
      ds.getRange("A7").getValue() + "   --->   " + ds.getRange("B7").getValue() + "\r" +
      ds.getRange("A8").getValue() + "   --->   " + ds.getRange("B8").getValue() + "\r" +
      "\rTanggal Request: " + tgl + "\rTimestamp: " + new Date();

    // Logger.log(array_sumber, array_isi);
  }
  
  GmailApp.sendEmail(email, subjek, body);
  Logger.log('Email terkirim ke: '+email);

}

function ambilSisaBulanan(){
 var url = 'aabbcc';   //Ganti dengan URL dari Google Sheet;
  var ss= SpreadsheetApp.openByUrl(url);
  var ds = ss.getSheetByName("Sheet2");
  var sisa = ds.getRange("B8").getValue();
  return sisa;
}
