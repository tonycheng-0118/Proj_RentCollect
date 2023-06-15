/////////////////////////////////////////
// Golbal Var
/////////////////////////////////////////
// Release vs Verify
// const CONST_TODAY_DATE = new Date(); 
// const SheetHandle      = SpreadsheetApp.openById("1mfZNNZLIzVbhaUwr88iEe8XUWP0qnHdmF8x2eVpXfVY"); // Proj_RentCollect
// const MODE_VERIFY      = false;
const CONST_TODAY_DATE = new Date("Wed May 17 2023 22:46:10 GMT+0800 (Taipei Standard Time)"); // for Proj_RentCollect_VERIFY, it is fixed.
const SheetHandle      = SpreadsheetApp.openById("1sSeMT7ZnQtuwSbRs0W-zqorA0MD_xRg0Fu8iPbg_YS8"); // Proj_RentCollect_VERIFY.
const MODE_VERIFY      = true;

// Sheet Name
const SheetImportName       = SheetHandle.getSheetByName('Import');
const SheetBankRecordName   = SheetHandle.getSheetByName('BankRecord');
const SheetBankRecordBKName = SheetHandle.getSheetByName('BankRecordBK');
const SheetContractName     = SheetHandle.getSheetByName('Contract');
const SheetTenantName       = SheetHandle.getSheetByName('Tenant');
const SheetUtilBillName     = SheetHandle.getSheetByName('UtilBill');
const SheetMiscCostName     = SheetHandle.getSheetByName('MiscCost');
const SheetPropertyName     = SheetHandle.getSheetByName('Property');
const SheetRptStatusName    = SheetHandle.getSheetByName('RptStatus');
const SheetRptEventName     = SheetHandle.getSheetByName('RptEvent');
const SheetREADMEName       = SheetHandle.getSheetByName('README');

const CONST_MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
const CONST_This_Account_Number = "000014853**1373*"; // The self account number shown in the Deposit Apply Record


function rentCollect_main() {
  Logger.log("This rentCollect_main");
  
  // to remove filter from sheet in case of not closing the filter.
  removeFilter();
  
  //
  rentCollect_parser();
  rentCollect_contract();
  rentCollect_report();

  //
  rentCollect_debug_print();
}

function removeFilter(){
  var sheetName_arr = [SheetImportName,SheetBankRecordName,SheetBankRecordBKName,SheetContractName,SheetTenantName,SheetUtilBillName,SheetMiscCostName,SheetPropertyName,SheetRptStatusName,SheetRptEventName];
  
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

function rentCollect_debug_print() {

  // locate ErrorLog row offset
  var tf = SheetREADMEName.createTextFinder("ErrorLog");
  var all = tf.findAll();
  if (all.length > 1 ) {var errMsg = `[rentCollect_debug_print] More then one ErrorLog location!`; reportErrMsg(errMsg);}
  if (all.length == 0) {var errMsg = `[rentCollect_debug_print] No ErrorLog location!`; reportErrMsg(errMsg);}
  
  // for (var i=0;i<all.length;i++) {
  //   Logger.log(`all[${i}].getA1Notation(): ${all[i].getA1Notation()}, all[${i}].getValue(): ${all[i].getValue()}, all[i].getRow(): ${all[i].getRow()}, all[i].getColumn(): ${all[i].getColumn()}`);
  // }
  
  var PosRowBase = all[all.length-1].getRow();
  var PosColBase = all[all.length-1].getColumn();
  if (SheetREADMEName.getLastRow() > PosRowBase) {
    SheetREADMEName.getRange(1+PosRowBase,PosColBase,SheetREADMEName.getLastRow()-PosRowBase,SheetREADMEName.getLastColumn()-PosColBase+1).clear();
  }
  
  // show status
  var logPos=0;
  if (errorLogCollection_arr.length > 0) {
    var text = `Error @ ${new Date()}`;
    SheetREADMEName.getRange(1+PosRowBase,PosColBase).setValue(text).setFontColor("#F08080"); // red
    logPos+=1;
    
    for (var i=0;i<errorLogCollection_arr.length;i++) {
      SheetREADMEName.getRange(logPos+1+PosRowBase,PosColBase).setValue(errorLogCollection_arr[i]).setBackground("#F08080"); // red
      logPos+=1;
    }
  } 
  
  if (warnLogCollection_arr.length > 0) {
    var text = `Warn @ ${new Date()}`;
    SheetREADMEName.getRange(1+PosRowBase,PosColBase).setValue(text).setFontColor("#FF8C00"); // yellow
    logPos+=1;
    
    for (var i=0;i<warnLogCollection_arr.length;i++) {
      SheetREADMEName.getRange(logPos+1+PosRowBase,PosColBase).setValue(warnLogCollection_arr[i]).setBackground("#FF8C00"); // yellow
      logPos+=1;
    }
  } 
  
  if (logPos == 0){
    var text = `PASS @ ${new Date()}`;
    SheetREADMEName.getRange(1+PosRowBase,PosColBase).setValue(text).setFontColor("#3CB371"); // green
  }

}


