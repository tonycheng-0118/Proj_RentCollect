function rentCollect_main() {
  
  //start
  var time_start = new Date();
  Logger.log(`rentCollect_main start @ ${Utilities.formatDate(time_start, "GMT+8", "HH:mm:ss")}`);
  
  // to remove filter from sheet in case of not closing the filter.
  removeFilter();

  // top setting
  mainCFG();
  
  // main
  rentCollect_parser();
  rentCollect_contract();
  rentCollect_report();

  // finish
  var time_finish = new Date();
  var time_exec = (time_finish.getTime() - time_start.getTime()) / 1000; // getTime() is ms
  Logger.log(`rentCollect_main finish @ ${Utilities.formatDate(time_finish, "GMT+8", "HH:mm:ss")}, time_exec(s): ${time_exec}`);
  if (1) {
    var info = `[rentCollect_main] time_exec(s): ${time_exec}`; reportInfoMsg(info);
  }

  // error log
  rentCollect_debug_print();

}


function mainCFG(){
  // getCFG
  for (i=0;i<CFG_Key_arr.length;i++){
    var cfg_key = CFG_Key_arr[i];
    var tf = SheetREADMEName.createTextFinder(cfg_key);
    var all = tf.findAll();
    if (all.length > 1 ) {var errMsg = `[getCFG] ${cfg_key} has more then one CFG location!`; reportErrMsg(errMsg);}
    if (all.length == 0) {var errMsg = `[getCFG] ${cfg_key} does not present!`; reportErrMsg(errMsg);}
    var posRowBase = all[all.length-1].getRow();
    var posColBase = all[all.length-1].getColumn();
    CFG_Val_obj[cfg_key] = SheetREADMEName.getRange(posRowBase+1,posColBase).getValues();
  }
  
  Object.keys(CFG_Val_obj).forEach (
    function(key) {
      Logger.log(`CFG: key:${key}, value:${CFG_Val_obj[key]}`);
    }
  )

  // set CONST
  CONST_TODAY_DATE.setSeconds(0);CONST_TODAY_DATE.setMinutes(0);CONST_TODAY_DATE.setHours(0);
}

function removeFilter(){
  var sheetName_arr = [SheetImportName,SheetBankRecordName,SheetBankRecordBKName,SheetContractName,SheetTenantName,SheetUtilBillName,SheetMiscCostName,SheetPropertyName,SheetRptStatusName,SheetRptEventName,SheetRptAnalysisName];
  
  sheetName_arr.forEach(
    function(sheetName){
      var filter = sheetName.getFilter();
      if (filter !== null) {
        var range = filter.getRange(); 
        var firstColumn = range.getColumn();
        var lastColumn = range.getLastColumn();
        // Logger.log(`name: ${sheetName.getName()}, firstColumn: ${firstColumn}, lastColumn: ${lastColumn}`);
        for (var i = firstColumn; i <= lastColumn; i++) {
          filter.removeColumnFilterCriteria(i);
        }
      }
    }
  )
}

var errorLogCollection_arr = new Array(); // to collect all all error string. format: [type] erro_msg
function reportErrMsg(errMsg){
  if (errMsg!=null) {
    var err = '[Error]' + errMsg;
    console.error(err);
    errorLogCollection_arr.push(err);
  }
}

var warnLogCollection_arr = new Array(); // to collect all all warn string. format: [type] warn_msg
function reportWarnMsg(errMsg){
  if (errMsg!=null) {
    var err = '[Warn]' + errMsg;
    console.warn(err);
    warnLogCollection_arr.push(err);
  }
}

var infoLogCollection_arr = new Array(); // to collect all all info string. format: [type] info_msg
function reportInfoMsg(errMsg){
  if (errMsg!=null) {
    var err = '[Info]' + errMsg;
    console.info(err);
    infoLogCollection_arr.push(err);
  }
}

var warnGenUtilBill_arr = new Array(); // to auto gen the potental UtilBill item. format: [type] UtilBill item
function reportWarnGenUtilBill(errMsg){
  if (errMsg!=null) {
    var err = '[WarnGenUtilBill]\t' + errMsg;
    console.warn(err);
    warnGenUtilBill_arr.push(err);
  }
}

var warnGenMiscCost_arr = new Array(); // to auto gen the potental MiscCost item. format: [type] MiscCost item
function reportWarnGenMiscCost(errMsg){
  if (errMsg!=null) {
    var err = '[WarnGenMiscCost]\t' + errMsg;
    console.warn(err);
    warnGenMiscCost_arr.push(err);
  }
}

function rentCollect_debug_print() {

  // locate ErrorLog row offset
  var tf = SheetREADMEName.createTextFinder("ErrorLog");
  var all = tf.findAll();
  if (all.length > 1 ) {var errMsg = `[rentCollect_debug_print] More then one ErrorLog location!`; reportErrMsg(errMsg);}
  if (all.length == 0) {var errMsg = `[rentCollect_debug_print] No ErrorLog location!`; reportErrMsg(errMsg);}
  
  // for (var i=0;i<all.length;i++) {
  //   Logger.log(`all[${i}].getA1Notation(): ${all[i].getA1Notation()}, all[${i}].getValue(): ${all[i].getValue()}, all[i].getRow(): ${all[i].getRow()}, all[i].getColumn(): ${all[i].getColumn()}`);
  // }
  
  var posRowBase = all[all.length-1].getRow();
  var posColBase = all[all.length-1].getColumn();
  if (SheetREADMEName.getLastRow() > posRowBase) {
    SheetREADMEName.getRange(1+posRowBase,posColBase,SheetREADMEName.getLastRow()-posRowBase+1,SheetREADMEName.getLastColumn()-posColBase+1).setNote(null).clear();
  }
  
  // show status
  var logPos=0;
  var text = `Updated date @ ${new Date()}`;
  SheetREADMEName.getRange(1+posRowBase,posColBase).setValue(text);
  logPos+=1;

  if (errorLogCollection_arr.length > 0) {
    for (var i=0;i<errorLogCollection_arr.length;i++) {
      SheetREADMEName.getRange(logPos+1+posRowBase,posColBase).setValue(errorLogCollection_arr[i]).setBackground("#F08080"); // red
      logPos+=1;
    }
  } 

  
  if (warnLogCollection_arr.length > 0) {
    // for UtilBill paste
    if (warnGenUtilBill_arr.length > 0) {
      var text = `WarnGenUtilBill List`;
      SheetREADMEName.getRange(logPos+1+posRowBase,posColBase).setValue(text).setBackground("#CBC3E3").setNote(`${warnGenUtilBill_arr}`); // Red Purple
      logPos+=1;
    }

    // for MiscCost paste
    if (warnGenMiscCost_arr.length > 0) {
      var text = `WarnGenMiscCost List`;
      SheetREADMEName.getRange(logPos+1+posRowBase,posColBase).setValue(text).setBackground("#CF9FFF").setNote(`${warnGenMiscCost_arr}`); // Light Violet
      logPos+=1;
    }
    
    for (var i=0;i<warnLogCollection_arr.length;i++) {
      SheetREADMEName.getRange(logPos+1+posRowBase,posColBase).setValue(warnLogCollection_arr[i]).setBackground("#FF8C00"); // yellow
      logPos+=1;
    }
  }

  if (logPos == 0){
    var text = `PASS @ ${new Date()}`;
    SheetREADMEName.getRange(logPos+1+posRowBase,posColBase).setValue(text).setFontColor("#3CB371"); // green
    logPos+=1;
  }

  if (infoLogCollection_arr.length > 0) {
    for (var i=0;i<infoLogCollection_arr.length;i++) {
      SheetREADMEName.getRange(logPos+1+posRowBase,posColBase).setValue(infoLogCollection_arr[i]).setBackground("#3CB371"); // green
      logPos+=1;
    }
  }

}