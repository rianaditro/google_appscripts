function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Scripts')
      .addItem('Buat Jadwal Konsul', 'runCreateLatestEventManual')
      .addItem('Hitung Margin', 'recalculateFormulas')
      .addToUi();
  
    Logger.log("onOpen: Collecting existing dates...");
    var existingDates = collectExistingDates();
    PropertiesService.getScriptProperties().setProperty('existingDates', JSON.stringify(existingDates));
    Logger.log("onOpen: Existing dates stored successfully.");
  }
  
  function collectExistingDates() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var existingDates = [];
  
    for (var i = 2; i < data.length; i++) {
      for (var j = 44; j <= 54; j++) {
        var dateValue = new Date(data[i][j]);
        if (!isNaN(dateValue.getTime())) {
          existingDates.push(dateValue);
        }
      }
    }
  
    Logger.log("collectExistingDates: Found " + existingDates.length + " existing dates.");
    return existingDates;
  }
  
  function runCreateLatestEventManual() {
    var stored = PropertiesService.getScriptProperties().getProperty('existingDates') || "[]";
    var existingDates = JSON.parse(stored).map(dateStr => new Date(dateStr));
    Logger.log("runCreateLatestEventManual: Loaded " + existingDates.length + " stored dates.");
    createLatestEvent(existingDates);
  }
  
  function createLatestEvent(existingDates) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var ui = SpreadsheetApp.getUi();
  
    for (var i = 2; i < data.length; i++) {
      var patientEmail = data[i][2];
      var doctorEmail = data[i][41];
      var status = data[i][43];
      
      if (status.toString().trim().toUpperCase() !== "BARU") {
        Logger.log("Row " + (i + 1) + ": event is already exist.");
        continue;
      }
  
      var scheduledEvent = [];
      for (var j = 44; j < 54; j++) {
        var dateValue = new Date(data[i][j]);
        if (!isNaN(dateValue.getTime())) {
          scheduledEvent.push(dateValue);
        }
      }
  
      var latestDate = findLatestDate(scheduledEvent);
      if (!latestDate) {
        Logger.log("Row " + (i + 1) + ": No valid date found.");
        ui.alert("❌ Error", "Baris " + (i + 1) + ": Format tanggal tidak valid!", ui.ButtonSet.OK);
        continue;
      }
  
      if (isDuplicateOrNearby(existingDates, latestDate)) {
        Logger.log("Row " + (i + 1) + ": Date " + latestDate + " is duplicate or too close.");
        ui.alert("❌ Konflik Jadwal", "Baris " + (i + 1) + ": Tanggal " + latestDate + " bentrok dengan jadwal yang ada.", ui.ButtonSet.OK);
        continue;
      }
  
      existingDates.push(latestDate);
      Logger.log("Row " + (i + 1) + ": Membuat acara pada " + latestDate);
      
      try {
        var event = CalendarApp.getDefaultCalendar().createEvent(
          "Konsultasi Medis",
          latestDate,
          new Date(latestDate.getTime() + 10 * 60 * 1000),
          { guests: patientEmail + "," + doctorEmail, sendInvites: true }
        );
        event.addEmailReminder(60);
        event.addPopupReminder(10);
  
        // Update status in column AR (index 43) to "TERJADWAL"
        sheet.getRange(i + 1, 44).setValue("TERJADWAL");

        Logger.log("Row " + (i + 1) + ": Event berhasil dibuat!");
        ui.alert("✅ Sukses!", "Jadwal berhasil dibuat untuk " + latestDate, ui.ButtonSet.OK);
      } catch (error) {
        Logger.log("ERROR: " + error.message);
        ui.alert("❌ Kesalahan!", "Gagal membuat acara: " + error.message, ui.ButtonSet.OK);
      }
    }
  }
  
  // Check if newDate is duplicate or within 14 minutes of any existing date
  function isDuplicateOrNearby(existingDates, newDate) {
    return existingDates.some(date => Math.abs(date.getTime() - newDate.getTime()) <= 14 * 60 * 1000);
  }
  
  // Returns the latest date from an array of dates
  function findLatestDate(dates) {
    if (dates.length === 0) return null;
    return dates.reduce((latest, date) => (new Date(date) > new Date(latest) ? date : latest), dates[0]);
  }
  
  // -------------------------------------------------------
  // Recalculation Function
  // -------------------------------------------------------
  
  // Recalculate formulas for rows where column AO (index 40) equals "HITUNG ULANG"
  function recalculateFormulas() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var numRows = data.length;
    var updatedRows = 0;
    
    for (var i = 2; i < numRows; i++) { // row 2 is header
      // Check if column AO (index 40) equals "HITUNG ULANG" (case-insensitive)
      if (data[i][40] && data[i][40].toString().trim().toUpperCase() === "HITUNG ULANG") {
        // Column M (index 12) = J (index 9) - K (index 10) - L (index 11)
        var valJ = parseFloat(data[i][9]) || 0;
        var valK = parseFloat(data[i][10]) || 0;
        var valL = parseFloat(data[i][11]) || 0;
        var newM = valJ - valK - valL;
        
        // Column T (index 19) = sum of O to S (indexes 14 to 18)
        var newT = 0;
        for (var col = 14; col <= 18; col++) {
          newT += parseFloat(data[i][col]) || 0;
        }
        
        // Column Z (index 25) = sum of U to Y (indexes 20 to 24)
        var newZ = 0;
        for (var col = 20; col <= 24; col++) {
          newZ += parseFloat(data[i][col]) || 0;
        }
        
        // Column AA is at index 26; get its value
        var valAA = parseFloat(data[i][26]) || 0;
        
        // Column AB (index 27) = T - Z - AA
        var newAB = newT - newZ - valAA;
        
        // Column AL (index 37) = H (index 7) + T + J
        var valH = parseFloat(data[i][7]) || 0;
        var newAL = valH + newT + valJ;
        
        // Column AM (index 38) = I (index 8) + U (index 20) + K (index 10)
        var valI = parseFloat(data[i][8]) || 0;
        var valU = parseFloat(data[i][20]) || 0;
        var newAM = valI + valU + valK;
        
        // Column AN (index 39) = AL - AM
        var newAN = newAL - newAM;
        
        // Write computed values back to the data array
        data[i][12] = newM;   // Column M
        data[i][19] = newT;   // Column T
        data[i][25] = newZ;   // Column Z
        data[i][27] = newAB;  // Column AB
        data[i][37] = newAL;  // Column AL
        data[i][38] = newAM;  // Column AM
        data[i][39] = newAN;  // Column AN
        
        //Update the "HITUNG ULANG" to "SELESAI" flag in column AO (index 40)
        data[i][40] = "SELESAI";
        
        updatedRows = i + 1;
      }
    }
    
    // Write the updated data back to the sheet in one batch
    dataRange.setValues(data);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert("✅ Perhitungan Margin Selesai", "Silahkan periksa kembali baris: " + updatedRows, ui.ButtonSet.OK);
  }
  
