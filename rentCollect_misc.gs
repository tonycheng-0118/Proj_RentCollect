function removeFilter(){ 
  var ss = SpreadsheetApp.getActive();
  var allsheets = ss.getSheets();
  allsheets.forEach(
    function(sheetName){
      var filter = sheetName.getFilter();
      if (filter !== null) {
        var range = filter.getRange(); 
        var firstColumn = range.getColumn();
        var lastColumn = range.getLastColumn();
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

  // to normalize CFG
  // set CONST
  CONST_TODAY_DATE.setSeconds(0);CONST_TODAY_DATE.setMinutes(0);CONST_TODAY_DATE.setHours(0);
  // to split the content
  VAR_WarnContract_arr = CFG_Val_obj["CFG_BankRecord_WarnContract"].toString().replace(/[\s|\n|\r|\t]/g,"").split(";");
}

function doLinePost(msg){
  const token =CFG_Val_obj["CFG_LinePostToken"].toString().replace(/[\s|\n|\r|\t]/g,"");//'UBcpzSiSgNnvRkoySrODkkjYswDkMjy0dzZ9UBSN9Dr';
  
  var message = "\n" + msg; // for 1st new line

  UrlFetchApp.fetch('https://notify-api.line.me/api/notify',{
    'headers':{'Authorization': 'Bearer ' + token,},
    'method':'post',
    'payload':{
    'message':message
    }
  });
}