/////////////////////////////////////////
// Golbal Var
/////////////////////////////////////////
// Release vs Verify
// const CONST_TODAY_DATE = new Date(); 
// const SheetHandle      = SpreadsheetApp.openById("1mfZNNZLIzVbhaUwr88iEe8XUWP0qnHdmF8x2eVpXfVY"); // Proj_RentCollect
const CONST_TODAY_DATE = new Date("Wed May 17 2023 22:46:10 GMT+0800 (Taipei Standard Time)"); // for Proj_RentCollect_VERIFY, it is fixed.
const SheetHandle      = SpreadsheetApp.openById("1sSeMT7ZnQtuwSbRs0W-zqorA0MD_xRg0Fu8iPbg_YS8"); // Proj_RentCollect_VERIFY.

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
  
  rentCollect_import();
  rentCollect_parser();
  rentCollect_contract();
  rentCollect_report();

  //
  rentCollect_debug_print();
}

var errorLogCollection_arr = new Array(); // to collect all all error string. format: [type] erro_msg
function reportErrMsg(errMsg){
  if (errMsg!=null) {
    var err = '[Error]' + errMsg;
    console.error(err);
    errorLogCollection_arr.push(err);
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

  if (errorLogCollection_arr.length==0) {
    var text = `PASS @ ${CONST_TODAY_DATE}`;
    SheetREADMEName.getRange(1+PosRowBase,PosColBase).setValue(text).setFontColor("#3CB371"); // green
  } else {
    for (var i=0;i<errorLogCollection_arr.length;i++) {
      SheetREADMEName.getRange(i+1+PosRowBase,PosColBase).setValue(errorLogCollection_arr[i]).setBackground("F08080"); // red
    }
  }

}


