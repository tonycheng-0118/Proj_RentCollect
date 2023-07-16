function rentCollect_import(depositRecordName, newRecordEnable, header_offset, tail_offset) {
  /////////////////////////////////////////
  // README
  /////////////////////////////////////////
  // Since the depositRecordName is overwritten once the sheet is exported.
  // It results in the different sheet ID for the latest one.
  // To address that, this function is to locate the sheet named "depositRecordName" and copy it into Import sheet.

  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const bankRecordBackUpEnable = (MODE_VERIFY == false); // if tesing merge SheetBankRecordName, turn this off to prevent from overriding SheetBankRecordBKName
  const depositRecordPreName = depositRecordName + "_pre";
  // const depositRecorColPosNote = 6; 


  /////////////////////////////////////////
  // Variable
  /////////////////////////////////////////
  var depositRecordFileHandle = DriveApp.getFilesByName(depositRecordName);
  var depositRecordSheetHandle;
  var depositRecordSheet;
  var isRecordExisted;

  /////////////////////////////////////////
  // Main
  /////////////////////////////////////////
  // Find the depositRecordName and depositRecordPreName in the current dir
  // Return the valid import status, if the Record is not existed, the post parser should be bypassed.
  if (depositRecordFileHandle.hasNext()) {
    var fileId = depositRecordFileHandle.next().getId();
    depositRecordSheetHandle = SpreadsheetApp.openById(fileId);
    depositRecordSheet = depositRecordSheetHandle.getSheets()[0]; // new depositRecordName sheet uploaded from somewhere
    isRecordExisted = true;
    Logger.log("[Info] %s ID is %s",depositRecordName,fileId)
    
    // copy depositRecordName to this SheetHandle
    if (SheetHandle.getSheetByName(depositRecordName)) {
      var sheet = SheetHandle.getSheetByName(depositRecordName);
      SheetHandle.deleteSheet(sheet); // dst sheet should not have this sheet name, just in case 
    }
    depositRecordSheet.copyTo(SheetHandle).setName(depositRecordName); // to solve cross sheet copy, copy the src sheet to dst sheet first

  } else {
    isRecordExisted = false;
    Logger.log("[Error] %s is not exiseted!",depositRecordName);
    // if (1) {var errMsg = `[rentCollect_import] ${depositRecordName} is not exiseted!`; reportErrMsg(errMsg);}
    if (1) {var warnMsg = `[rentCollect_import] ${depositRecordName} is not exiseted!`; reportWarnMsg(warnMsg);}
  }
  
  // import integrity check
  if (isRecordExisted) {
    var sts_integrity = false;
    var sheetTMPImportName    = SheetHandle.getSheetByName(depositRecordName);
    var sheetTMPImportPreName = SheetHandle.getSheetByName(depositRecordPreName);

    if (sheetTMPImportPreName){
      // check the overlap region bwt ImportPre and Import, should be large enough to ensure the Import quality
      Logger.log(`[Info] ${depositRecordPreName} exised, start to import integrity check.`);
      sts_integrity = chkImportIntegrity(sheetTMPImportName,sheetTMPImportPreName,header_offset,tail_offset);
    }
    else { // for 1st time case
      if (newRecordEnable){
        Logger.log(`[Info] ${depositRecordPreName} not existed, copy ${depositRecordName} to it.`);
        sheetTMPImportName.copyTo(SheetHandle).setName(depositRecordPreName);
        sts_integrity = true;
      }
      else {
        if (1) {var errMsg = `[rentCollect_import] For ${depositRecordName}, fail import integrity check, not enable new Record.`; reportErrMsg(errMsg);}
        sts_integrity = false;
      }
    }

    if (sts_integrity) {
      Logger.log(`[Info] import integrity check passed, start to copy ${depositRecordPreName} to SheetImportName.`);
      
      // backup BankRecord to BankRecordBK for in case
      if (bankRecordBackUpEnable) {
        Logger.log(`[Info] start to backup to SheetBankRecordBKName.`)
        SheetBankRecordBKName.getDataRange().clearContent();
        SheetBankRecordName.getDataRange().copyTo(SheetBankRecordBKName.getRange('A1'));
      }

      // copy from depositRecordName to ImportPre
      // SheetImportPreName.getRange('A:G').clearContent();
      // sheetTMP.getRange('A:G').copyTo(SheetImportPreName.getRange('A1'));
      // deleteEmptyRows(SheetImportPreName);
      
      // copy from depositRecordName to Import
      SheetImportName.getDataRange().clearContent();
      sheetTMPImportName.getDataRange().copyTo(SheetImportName.getRange('A1')); // ref to https://is.gd/EFpjSN

      // chagne depositRecordName to depositRecordPreName
      SheetHandle.deleteSheet(SheetHandle.getSheetByName(depositRecordPreName));
      sheetTMPImportName.setName(depositRecordPreName);
      
    } else {
      SheetHandle.deleteSheet(SheetHandle.getSheetByName(depositRecordName));
      Logger.log("[Error] Fail import integrity check, no touch to SheetImportName.");
      if (1) {var errMsg = `[rentCollect_import] For ${depositRecordName}, fail import integrity check, will not touch to SheetImportName.`; reportErrMsg(errMsg);}
    }
  }

  /////////////////////////////////////////
  // Post processing the content
  /////////////////////////////////////////
  // rid of the empty row
  deleteEmptyRows(SheetImportName);
  
  // rid of unused info
  // replaceRowContents(depositRecorColPosNote,THIS_ACCOUNT_NUMBER,"");

  return (isRecordExisted & sts_integrity);


}

function chkImportIntegrity(sheet_src,sheet_dst,header_offset,tail_offset){
  /////////////////////////////////////////
  // The idea of integrity is to check the consecutive match item in dst.
  /////////////////////////////////////////
  var src_arr = new Array();
  var dst_arr = new Array();
  var src = sheet_src.getDataRange().getValues();
  var dst = sheet_dst.getDataRange().getValues();
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
  var match_end   = false;
  for (var i =header_offset;i<dst_arr.length-tail_offset;i++){ // i=0 is header
    for (var j=j_cur+1;j<src_arr.length-tail_offset;j++){
      // Logger.log(`KKK: dst[${i}]: ${dst_arr[i]},\n     src[${j}]: ${src_arr[j]}`);
      if (dst_arr[i] == src_arr[j]) {
        match_start = true;
        if (match_end == false) pass = true;
        match_cnt++;
        j_cur = j;
        // Logger.log(`YES. Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`);
        break;
      }
      else {
        // Logger.log(`No,  Matched count: ${match_cnt},\n dst[${i}]: ${dst_arr[i]},\n src[${j}]: ${src_arr[j]}`);
        if (match_start) {
          if (match_end == false) pass = false;
          match_end = true;
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
