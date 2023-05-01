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
  
  // copy DEPOSIT_APPLY_RECORD to this SheetHandle
  if (SheetHandle.getSheetByName(depositRecordName)) {
    var sheet = SheetHandle.getSheetByName(depositRecordName);
    SheetHandle.deleteSheet(sheet); // dst sheet should not have this sheet name, just in case 
  }
  depositRecordSheet.copyTo(SheetHandle).setName(depositRecordName); // to solve cross sheet copy, copy the src sheet to dst sheet first
  
  if (SheetImportName) {
    var sheetTMP = SheetHandle.getSheetByName(depositRecordName);
    
    // backup BankRecord to BankRecordBK for in case
    SheetBankRecordBKName.getDataRange().clearContent();
    SheetBankRecordName.getDataRange().copyTo(SheetBankRecordBKName.getRange('A1'));

    // copy from DEPOSIT_APPLY_RECORD to ImportPre
    // SheetImportPreName.getRange('A:G').clearContent();
    // sheetTMP.getRange('A:G').copyTo(SheetImportPreName.getRange('A1'));
    // deleteEmptyRows(SheetImportPreName);

    // check the overlap region bwt ImportPre and Import, should be large enough to ensure the Import quality
    var sts_integrity = chkImportIntegrity(sheetTMP);
    
    // copy from DEPOSIT_APPLY_RECORD to Import
    if (sts_integrity) {
      SheetImportName.getRange('A:G').clearContent();
      sheetTMP.getRange('A:G').copyTo(SheetImportName.getRange('A1')); // ref to https://is.gd/EFpjSN
    }
    
  } else {
    // only for the 1st run
    Logger.log("[Error] Should have import sheet.");
  }

  // delete the DEPOSIT_APPLY_RECORD
  SheetHandle.deleteSheet(SheetHandle.getSheetByName(depositRecordName));


  /////////////////////////////////////////
  // Post processing the content
  /////////////////////////////////////////
  // rid of the empty row
  deleteEmptyRows(SheetImportName);
  
  // rid of unused info
  // replaceRowContents(depositRecorColPosNote,THIS_ACCOUNT_NUMBER,"");


}

function chkImportIntegrity(sheet_src){
  var src_arr = new Array();
  var dst_arr = new Array();
  // var src = SheetImportPreName.getDataRange().getValues();
  var src = sheet_src.getDataRange().getValues();
  var dst = SheetImportName.getDataRange().getValues();
  var pass = false;

  for(var i=0;i<src.length;i++){
    if(src[i].join().replace(/,/g,'')!=''){ 
      src_arr.push(src[i].join().toString().replace(/[\s|\n|\r|\t]/g,""))
    }
  }

  for(var i=0;i<dst.length;i++){
    if(dst[i].join().replace(/,/g,'')!=''){ 
      dst_arr.push(dst[i].join().toString().replace(/[\s|\n|\r|\t]/g,""))
    }
  }

  var j_cur=0;
  var match_cnt=0;
  var match_start = false;
  for (var i =0;i<dst_arr.length;i++){
    for (var j=j_cur+1;j<src_arr.length;j++){
      if (dst_arr[i] == src_arr[j]) {
        match_start = true;
        pass = true;
        match_cnt++;
        j_cur = j;
        // Logger.log(`YES. Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`);
        break;
      }
      else {
        // Logger.log(`No,  Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`);
        if (match_start) {
          pass = false;
          Logger.log(`Failed,  Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`);
          if (1) {var errMsg = `[chkImportIntegrity] Should be consecutive matched. Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`; reportErrMsg(errMsg);}
        }
      }
    }
    // Logger.log(`Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j_cur}]: ${src_arr[j_cur]}`);
  }

  if (match_start==false) {var errMsg = `[chkImportIntegrity] No matched at all. Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`; reportErrMsg(errMsg);}

  return pass;

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

function deleteEmptyRows(sh){ 
  // var sh = SheetHandle.getSheetByName("Import");
  // var sh = SpreadsheetApp.getActiveSheet();
  // Logger.log(sh.getName());
  var data = sh.getDataRange().getValues();
  var targetData = new Array();
  for(n=0;n<data.length;++n){
    if(data[n].join().replace(/,/g,'')!=''){ 
      targetData.push(data[n])
    }
    else {
      // Logger.log("Got an empty row at %d!", (n+1)); // row start at 1
    }
    // Logger.log(data[n].join().replace(/,/g,''));
  }
  sh.getDataRange().clear();
  sh.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);
}
