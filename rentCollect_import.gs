function rentCollect_import() {
  /////////////////////////////////////////
  // README
  /////////////////////////////////////////
  // Since the DEPOSIT_APPLY_RECORD is overwritten once the sheet is exported.
  // It results in the different sheet ID for the latest one.
  // To address that, this function is to locate the sheet named "DEPOSIT_APPLY_RECORD" and copy it into Import sheet.

  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const depositRecordName = "DEPOSIT_APPLY_RECORD";
  const depositRecorColPosNote = 6; 

  
  /////////////////////////////////////////
  // Variable
  /////////////////////////////////////////
  var depositRecordFileHandle = DriveApp.getFilesByName(depositRecordName);
  var depositRecordSheetHandle;
  var depositRecordSheet;

  /////////////////////////////////////////
  // Main
  /////////////////////////////////////////
  if (depositRecordFileHandle.hasNext()) {
    var fileId = depositRecordFileHandle.next().getId();
    depositRecordSheetHandle = SpreadsheetApp.openById(fileId);
    depositRecordSheet = depositRecordSheetHandle.getSheets()[0]; // new DEPOSIT_APPLY_RECORD sheet uploaded from somewhere
    Logger.log("[Info] %s ID is %s",depositRecordName,fileId)
  } else {
    Logger.log("[Error] %s is not exiseted!",depositRecordName);
  }

  if (SheetHandle.getSheetByName(depositRecordName)) {
    var sheet = SheetHandle.getSheetByName(depositRecordName);
    SheetHandle.deleteSheet(sheet); // dst sheet should not have this sheet name, just in case 
  }
  depositRecordSheet.copyTo(SheetHandle).setName(depositRecordName); // to solve cross sheet copy, copy the src sheet to dst sheet first
  
  if (SheetHandle.getSheetByName("Import")) {
    var sheet = SheetHandle.getSheetByName("Import");
    var cleanRagne = sheet.getRange('A:G');
    var sheetTMP = SheetHandle.getSheetByName(depositRecordName);
    var contentRange = sheetTMP.getRange('A:G');
    
    cleanRagne.clearContent();
    contentRange.copyTo(sheet.getRange('A1')); // ref to https://is.gd/EFpjSN
    SheetHandle.deleteSheet(sheetTMP);
  } else {
    // only for the 1st run
    Logger.log("[Error] Should have import sheet.");
  }


  /////////////////////////////////////////
  // Post processing the content
  /////////////////////////////////////////
  // rid of the empty row
  deleteEmptyRows();
  
  // rid of unused info
  // replaceRowContents(depositRecorColPosNote,THIS_ACCOUNT_NUMBER,"");


}

function replaceRowContents(columnPos, searchPTN, replacePTN){
  var sh = SheetHandle.getSheetByName("Import");
  // var sh = SpreadsheetApp.getActiveSheet(); // Will be the 1st sheet
  Logger.log(sh.getName());
  var data = sh.getDataRange().getValues();
  var targetData = new Array();
  for(n=0;n<data.length;++n){
    data[n][columnPos] = data[n][columnPos].replace(searchPTN,replacePTN);
    targetData.push(data[n]);
  }
  sh.getDataRange().clear();
  sh.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);
}

function deleteEmptyRows(){ 
  var sh = SheetHandle.getSheetByName("Import");
  // var sh = SpreadsheetApp.getActiveSheet();
  Logger.log(sh.getName());
  var data = sh.getDataRange().getValues();
  var targetData = new Array();
  for(n=0;n<data.length;++n){
    if(data[n].join().replace(/,/g,'')!=''){ 
      targetData.push(data[n])
    }
    else {
      Logger.log("Got an empty row at %d!", (n+1)); // row start at 1
    }
    // Logger.log(data[n].join().replace(/,/g,''));
  }
  sh.getDataRange().clear();
  sh.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);
}
