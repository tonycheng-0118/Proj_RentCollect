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

// Seperated SheetHandle
const SheetHandle_MinSheng35 = SpreadsheetApp.openById("1k2yNsmI-pTXqUsgW6On4ulyssaQUTmrxnMsYNYrYUl8"); // Proj_RentCollect_MinSheng35.

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
const SheetRptAnalysisName  = SheetHandle.getSheetByName('RptAnalysis');
const SheetLineMsgName      = SheetHandle.getSheetByName('LineMsg');
const SheetREADMEName       = SheetHandle.getSheetByName('README');

const CONST_MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
const CONST_This_Account_Number = "000014853**1373*"; // The self account number shown in the Deposit Apply Record
const CONST_SuperFeatureDate = new Date("Jan 1 3000 00:00:00 GMT+0800 (Taipei Standard Time)");
const CONST_SuperAncentDate = new Date("Jan 1 1000 00:00:00 GMT+0800 (Taipei Standard Time)");

const CONST_LinePostRentRecordWeek = 1;
const CONST_LinePostVacancyWeek = 6;
const CONST_LinePostOverdueWeek = 6;
const CONST_LinePostDayRestWeek = 6;

const CFG_Key_arr = ["CFG_BankRecordSearch_FromDateMargin","CFG_BankRecordSearch_ToDateMargin","CFG_ReportArrearMargin","CFG_BankRecordCheck_AmountMargin","CFG_BankRecordCheck_Details","CFG_MonthAccRent_NUM","CFG_ReportEvent_ShowEndContract","CFG_BankRecord_WarnContract","CFG_LinePostToken","CFG_NewRecordImport","CFG_NewProxyRecordImportNum","CFG_LineMsg_FromDate","CFG_LinePost_Enable"];
var CFG_Val_obj = new Object();

// Line userId
const CONST_LINE_USERID_TONY      = `Uf45e2cc323ccd4edcb12736d15c6da99`; // to promote this as a array in CFG
const CONST_LINE_USERID_ANGENT_0  = `Uf0ccd827ea24e45c6ede652cd49eb279`;

// Array
const CONST_CTBC_actWithdrawList = ["現金","中信卡","委代扣綜所稅","信託"]; // A list to collect all of the known withdraw action apart from Tranfer
const CONST_CTBC_actDepositList  = ["現金", "利息", "委代入","現金存款機","委代入補助款","委代入貨物稅","信託"]; // A list to collect all of the known deposit action apart from Tranfer
const CONST_KTB_actList = ["現金","現金D","現金E","本交A","利息D","退票D","稅款D","電費D","水費D"];
const CONST_MinSheng35_ID = ["2A","2B","2C","4A","4B","4C","5A","5B","5C","6A","6B","6C","7D","7E","8D","8E"];
const GLB_Tenant_AccountName_Pos = 0;
const GLB_Tenant_Account_Pos = 1;

var GLB_Tenant_obj      = new Object();
var GLB_Import_arr      = new Array();
var GLB_BankRecord_arr  = new Array();
var GLB_Contract_arr    = new Array();
var GLB_UtilBill_arr    = new Array();
var GLB_MiscCost_arr    = new Array();
var GLB_Property_arr    = new Array();
var GLB_RptStatus_arr   = new Array();
var GLB_RptEvent_arr    = new Array();
var GLB_LineMsg_arr     = new Array();

// var
var VAR_rptEventItemNo = 0;
var VAR_CompareBankRecord_arr = new Array();
var VAR_WarnContract_arr = new Array();
var VAR_ContractNoList_arr = new Array();
var VAR_isNewRecordImport = false;

/////////////////////////////////////////
// Class
/////////////////////////////////////////
class itemLineMsg {

  constructor (){
    this.itemNo           = 0;
    this.date             = new Date();
    this.timestamp        = "";
    this.userId           = ""; 
    this.userName         = "";
    this.type             = "";
    this.content          = "";
    this.note             = "";
    this.checkSum         = "";
    
    this.itemPack         = null;
    this.itemPackMaxLen   = 9;
        
    // if (this.itemPack.length == this.itemPackMaxLen) {
    //   this.contractNo             = item[6];
    //   this.tenantName             = item[7];
    //   this.validContract          = item[8];
    // }
    // else if (this.itemPack.length > this.itemPackMaxLen) {
    //   if (1) {var errMsg = `[itemUtilBill] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    // }
  };

  extract (item) {
    this.itemNo           = item[0];
    this.date             = item[1];
    this.timestamp        = item[2];
    this.userId           = item[3]; 
    this.userName         = item[4];
    this.type             = item[5];
    this.content          = item[6];
    this.note             = item[7];
    this.checkSum         = item[8];
    
    this.itemPack         = item;
  };

  compare (item){
    var that = new itemLineMsg();
    that.extract(item);
    // Logger.log(`Compare\n this: ${this.show()}\n that: ${that.show()}`);
    if (this.checkSum == that.checkSum) {
      return true;
    } else {
      return false
    }
  }

  show(){
    var text = `itemLineMsg: \n(itemNo=${this.itemNo},date=${this.date},timestamp=${this.timestamp},userId=${this.userId},userName=${this.userName},type=${this.type},content=${this.content},note=${this.note},checkSum=${this.checkSum})`;
    return text;
  };

}

class itemProperty {

  constructor (item){
    this.itemNo                 = item[0];
    this.rentProperty           = item[1];
    this.note                   = item[2];
    this.occupied               = null;
    this.dayRest                = null;
    this.curRent                = null;
    this.contractNo             = null;
    this.tenantName             = null;
    this.validContract          = null;
    this.ColPos_Occupied        = 4;
    this.ColPos_DayRest         = 5;
    this.ColPos_CurRent         = 6;
    this.ColPos_ContractNo      = 7;
    this.ColPos_TenantName      = 8;
    this.ColPos_ValidContract   = 9;

    this.itemPack               = item;
    this.itemPackMaxLen         = 9;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.occupied              = item[3];
      this.dayRest               = item[4];
      this.curRent               = item[5];
      this.contractNo            = item[6];
      this.tenantName            = item[7];
      this.validContract         = item[8];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemProperty] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.occupied                 = upd[0];
    this.dayRest                  = upd[1];
    this.curRent                  = upd[2];
    this.contractNo               = upd[3];
    this.tenantName               = upd[4];
    this.validContract            = upd[5];
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  compare (item){
    var that = new itemProperty(item);
    // other.show();
    if (
      (this.rentProperty         == that.rentProperty)
    ) {
      Logger.log(`AAA: ${this.show()}`);
      Logger.log(`BBB: ${that.show()}`);
      return true;
    } else return false
  };

  show(){
    var text = `itemProperty: \n(itemNo=${this.itemNo},rentProperty=${this.rentProperty},note=${this.note},occupied=${this.occupied},dayRest=${this.dayRest},curRent=${this.curRent},contractNo=${this.contractNo},validContract=${this.validContract})`;
    return text;
  };

}

class itemUtilBill {

  constructor (item){
    this.itemNo                 = item[0];
    this.date                   = item[1];
    this.rentProperty           = item[2]; 
    this.amount                 = item[3];
    this.note                   = item[4];
    this.contractOverrid        = item[5];
    this.contractNo             = null;
    this.tenantName             = null;
    this.validContract          = null;
    this.ColPos_ContractNo      = 7;
    this.ColPos_TenantName      = 8;
    this.ColPos_ValidContract   = 9;
    
    this.itemPack               = item;
    this.itemPackMaxLen         = 9;
        
    if (this.itemPack.length == this.itemPackMaxLen) {
      this.contractNo             = item[6];
      this.tenantName             = item[7];
      this.validContract          = item[8];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemUtilBill] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.contractNo             = upd[0];
    this.tenantName             = upd[1];
    this.validContract          = upd[2];
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  };

  compare (item){
    var that = new itemUtilBill(item);
    // other.show();
    if (
      (this.date.toString()         == that.date.toString()) &&
      (this.rentProperty            == that.rentProperty) && 
      (this.amount                  == that.amount) &&
      (this.note                    == that.note)
    ) {
      Logger.log(`AAA: ${this.show()}`);
      Logger.log(`BBB: ${that.show()}`);
      return true;
    } else return false
  }

  show(){
    var text = `itemUtilBill: \n(itemNo=${this.itemNo},date=${this.date},rentProperty=${this.rentProperty},amount=${this.amount},note=${this.note},contractOverrid=${this.contractOverrid},contractNo=${this.contractNo},tenantName=${this.tenantName},validContract=${this.validContract})`;
    // Logger.log(text);
    return text;
  };

}

class itemMiscCost {

  constructor (item){
    this.itemNo           = item[0];
    this.date             = item[1];
    this.rentProperty     = item[2]; 
    this.amount           = item[3];
    this.type             = item[4];
    this.note             = item[5];
    this.contractOverrid  = item[6];
    this.contractNo       = null;
    this.tenantName       = null;
    this.validContract    = null;
    
    this.ColPos_ContractNo      = 8;
    this.ColPos_TenantName      = 9;
    this.ColPos_ValidContract   = 10;

    this.MiscType_Charge_Fee    = `0.Charge_Fee`;
    this.MiscType_CashRent      = `1.Cash_Rent`;
    this.MiscType_SubRent       = `2.Sub_Rent`;
    this.MiscType_Commission    = `3.Commission`
    this.MiscType_Repare_Fee    = `5.Repare_Fee`;
    this.MiscType_Refund        = `6.Refund`;
    
    this.itemPack               = item;
    this.itemPackMaxLen         = 10;
    
    if (this.itemPack.length == this.itemPackMaxLen) {
      this.contractNo             = item[7];
      this.tenantName             = item[8];
      this.validContract          = item[9];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemMiscCost] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.contractNo      = upd[0];
    this.tenantName      = upd[1];
    this.validContract   = upd[2];
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  };

  compare (item){
    var that = new itemMiscCost(item);
    // other.show();
    if (
      (this.date.toString()         == that.date.toString()) &&
      (this.rentProperty            == that.rentProperty) && 
      (this.amount                  == that.amount) &&
      (this.type                    == that.type) &&
      (this.note                    == that.note)
    ) {
      Logger.log(`AAA: ${this.show()}`);
      Logger.log(`BBB: ${that.show()}`);
      return true;
    } else return false
  }

  expect_misc (){ // from expected payment perspective
    if      (this.type == this.MiscType_Charge_Fee) return ( 1 * this.amount); // extra tenant payment apaprt from rent.
    else if (this.type == this.MiscType_CashRent)   return (-1 * this.amount); // cash paid by the tenant
    else if (this.type == this.MiscType_SubRent)    return (-1 * this.amount); // rent deduction, the expense for other purpose.
    else if (this.type == this.MiscType_Commission) return (-1 * this.amount); // payment to the agent.
    else if (this.type == this.MiscType_Repare_Fee) return ( 0 * this.amount); // absorbed by preperty owner
    else if (this.type == this.MiscType_Refund)     return (-1 * this.amount); // return back to tenant
    else {var errMsg = `[rentCollect_contract] Type of MiscNo: ${this.itemNo} is invalid!`; reportErrMsg(errMsg);}
  }

  balance_misc (){ // from property balance perspective
    if      (this.type == this.MiscType_Charge_Fee) return ( 0 * this.amount); // will be covered by bank record
    else if (this.type == this.MiscType_CashRent)   return ( 1 * this.amount); // cash paid by the tenant
    else if (this.type == this.MiscType_SubRent)    return ( 1 * this.amount); // rent deduction, is sub for other usage, still income.
    else if (this.type == this.MiscType_Commission) return (-1 * this.amount); // payment to the agent.
    else if (this.type == this.MiscType_Repare_Fee) return (-1 * this.amount); // expect to be absorbed by preperty owner
    else if (this.type == this.MiscType_Refund)     return (-1 * this.amount); // return back to tenant
    else {var errMsg = `[rentCollect_contract] Type of MiscNo: ${this.itemNo} is invalid!`; reportErrMsg(errMsg);}
  }

  show(){
    var text = `itemMiscCost: \n(itemNo=${this.itemNo},date=${this.date},rentProperty=${this.rentProperty},amount=${this.amount},type=${this.type},note=${this.note},contractOverrid=${this.contractOverrid},contractNo=${this.contractNo},tenantName=${this.tenantName},validContract=${this.validContract})`;
    // Logger.log(text);
    return text;
  };

}

class itemBankRecord {

  constructor (item){
    this.itemNo                 = item[0];
    this.date                   = item[1];
    this.action                 = item[2]; 
    this.amount                 = item[3];
    this.balance                = item[4];
    this.fromAccountName        = item[5];
    this.fromAccount            = item[6];
    this.toAccountName          = item[7];
    this.toAccount              = item[8];
    this.mergeDate              = item[9];
    this.recordCheck            = item[10];
    this.recordNote             = item[11];
    this.contractOverrid        = item[12];
    this.contractNo             = item[13];
    this.rentProperty           = item[14];
    this.lineMsg                = item[15];
    this.ColPos_RecordCheck     = 11;
    this.ColPos_RecordNote      = 12;
    this.ColPos_ContractNo      = 14;
    this.ColPos_rentProperty    = 15;
    this.ColPos_lineMsg         = 16;
    
    this.itemPack               = item;
    this.itemPackMaxLen         = 16;

    if (this.itemPack.length != this.itemPackMaxLen) {
      var errMsg = `[itemBankRecord] Incorrect itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);
    }
  };

  update (upd, flag="") {
    if (flag=="lineMsg") {
      if (upd.length!=1) {var errMsg = `[itemBankRecord] lineMsg can only be updated along. @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
      this.lineMsg                = upd[0];
      this.itemPack[this.ColPos_lineMsg-1]      = upd[0];
    } else if (flag=="contractNo_rentProperty") {
      if (upd.length!=2) {var errMsg = `[itemBankRecord] contractNo and rentProperty can only be updated along. @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
      this.contractNo             = upd[0];
      this.rentProperty           = upd[1];
      this.itemPack[this.ColPos_ContractNo-1]   = upd[0];
      this.itemPack[this.ColPos_rentProperty-1] = upd[1];
    } else if (flag=="recordCheck") {
      if (upd.length!=1) {var errMsg = `[itemBankRecord] recordCheck can only be updated along. @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
      this.recordCheck            = upd[0];
      this.itemPack[this.ColPos_RecordCheck-1]  = upd[0];
    } else {
      {var errMsg = `[itemBankRecord] At least one flag for update. @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  compare (item){
    var that = new itemBankRecord(item);
    // other.show();
    if (
      (this.date.toString()         == that.date.toString()) &&
      (this.action                  == that.action) && 
      (this.amount                  == that.amount) &&
      // (this.fromAccountName         == that.fromAccountName) &&
      // (this.fromAccount             == that.fromAccount) &&
      // (this.toAccountName           == that.toAccountName) &&
      // (this.toAccount               == that.toAccount) &&
      (this.balance                 == that.balance)
    ) {
      Logger.log(`AAA: ${this.show()}`);
      Logger.log(`BBB: ${that.show()}`);
      return true;
    } else return false
  }

  show(){
    var text = `itemBankRecord: \n(itemNo=${this.itemNo},date=${this.date},action=${this.action},amount=${this.amount},balance=${this.balance},fromAccountName=${this.fromAccountName},fromAccount=${this.fromAccount},toAccountName=${this.toAccountName},toAccount=${this.toAccount},mergeDate=${this.mergeDate},recordCheck=${this.recordCheck},recordNote=${this.recordNote},contractOverrid=${this.contractOverrid},contractNo=${this.contractNo},rentProperty=${this.rentProperty})`;
    // Logger.log(text);
    return text;
  };

}

class itemContract {
  constructor (item){
    this.itemNo                   = item[0];
    this.fromDate                 = item[1];
    this.toDate                   = item[2]; 
    this.rentProperty             = item[3];
    this.deposit                  = item[4];
    this.amount                   = item[5];
    this.period                   = item[6];
    this.tenantName               = item[7];
    this.tenantAccountName_regex  = item[8];
    this.tenantAccount_arr        = item[9];
    this.toAccountName            = item[10];
    this.toAccount                = item[11];
    this.endContract              = item[12];
    this.note                     = item[13];
    this.fileLink                 = item[14];
    this.proxy                    = item[15];
    this.org                      = item[16];
    this.endDate                  = item[17];
    this.validContract            = null;
    this.rentArrear               = null;
    this.dayRest                  = null;
    
    this.ColPos_ItemNo            = 1;
    this.ColPos_Deposit           = 5;
    this.ColPos_EndDate           = 16;
    this.ColPos_ValidContract     = 17;
    this.ColPos_RentArrear        = 18;
    this.ColPos_DayRest           = 19;
    
    this.isProxying               = (item[15].toString().replace(/[\s|\n|\r|\t]/g,"") != "PROXIED") && (item[15].toString().replace(/[\s|\n|\r|\t]/g,"") != "");
    this.isProxied                = (item[15].toString().replace(/[\s|\n|\r|\t]/g,"") == "PROXIED");
    this.proxied_array            = this.unpack_proxied_array(item[15]);
    
    this.itemPack                 = item;
    this.itemPackMaxLen           = 21;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.validContract          = item[18];
      this.rentArrear             = item[19];
      this.dayRest                = item[20];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemContract] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  unpack_proxied_array(arr){
    var unpack_arr = arr.replace(/[\s|\n|\r|\t]/g,"").split(";");
    var return_arr = [];
    for (var i=0;i<unpack_arr.length;i++) {
      var regExp = new RegExp("([0-9]+)-([0-9]+)","gi");
      if (unpack_arr[i].match(regExp) != null) {
        var unpack = unpack_arr[i].split("-").toSorted();
        var start = unpack[0];
        var end   = unpack[1];
        for (var j=start;j<=end;j++){
          return_arr.push(j);
        }
      } else {
        return_arr.push(unpack_arr[i]);
      }
    }
    return return_arr;
  }

  update (upd) {
    this.endDate                  = upd[0];
    this.validContract            = upd[1];
    this.rentArrear               = upd[2];
    this.dayRest                  = upd[3];
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  compare (item){
    var that = new itemContract(item);
    // that.show();
    if (
      (this.itemNo.toString()       == that.itemNo.toString()) &&
      (this.fromDate.toString()     == that.fromDate.toString()) &&
      (this.toDate.toString()       == that.toDate.toString()) && 
      (this.rentProperty            == that.rentProperty) &&
      (this.tenantName              == that.tenantName) &&
      (this.tenantAccountName_regex == that.tenantAccountName_regex) &&
      (this.tenantAccount_arr.join() == that.tenantAccount_arr.join())
    ) return true
    else return false
  }

  show(){
    var text = `itemContract: \n(contractNo=${this.itemNo},fromDate=${this.fromDate},toDate=${this.toDate},rentProperty=${this.rentProperty},deposit=${this.deposit},amount=${this.amount},period=${this.period},tenantName=${this.tenantName},tenantAccountName_regex=${this.tenantAccountName_regex},tenantAccount_arr=${this.tenantAccount_arr},toAccountName=${this.toAccountName},toAccount=${this.toAccount},endContract=${this.endContract},note=${this.note},fileLink=${this.fileLink},proxy=${this.proxy},org=${this.org},endDate=${this.endDate},validContract=${this.validContract},rentArrear=${this.rentArrear},dayRest=${this.dayRest})`;
    // Logger.log(text);
    return text;
  };

}

class itemRptAnalysis {
  constructor (item) {
    this.itemNo;
    this.propertyGroup;
    this.propertyExclude;
    this.monthAccRent;
    
    this.ColPos_ItemNo            = 1;
    this.ColPos_PropertyGroupName = 2;
    this.ColPos_PropertyGroup     = 3;
    this.ColPos_PropertyExclude   = 4;
    this.ColPos_MonthAccRent      = 5;

    this.itemPack               = item;
    this.itemPackMaxLen         = 5;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.itemNo             = item[0];
      this.propertyGroupName  = item[1];
      this.propertyGroup      = item[2];
      this.propertyExclude    = item[3];
      this.monthAccRent       = item[4];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemRptAnalysis] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.itemNo             = upd[0];
    this.propertyGroupName  = upd[1];
    this.propertyGroup      = upd[2];
    this.propertyExclude    = upd[3];
    this.monthAccRent       = upd[4];
    
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  show(){
    var text = `itemRptStatus: \n(itemNo=${this.itemNo}, propertyGroupName=${this.propertyGroupName}, propertyGroup=${this.propertyGroup}, propertyExclude=${this.propertyExclude}, monthAccRent=${this.monthAccRent})`;
    return text;
  };

}

class itemRptStatus {
  constructor (item) {
    this.itemNo;
    this.rentProperty;
    this.tenantName;
    this.rentArrear;
    this.deposit;
    this.curRent;
    this.dayRest;
    this.contractNo;
    this.accBalance;
    this.status;
    this.note;
    this.occupied;
    this.validContract;

    this.ColPos_RentProperty    = 2;
    this.ColPos_TenantName      = 3;
    this.ColPos_RentArrear      = 4;
    this.ColPos_Deposit         = 5;
    this.ColPos_CurRent         = 6;
    this.ColPos_DayRest         = 7;
    this.ColPos_ContractNo      = 8;
    this.ColPos_AccBalance      = 9;
    this.ColPos_Status          = 10;
    this.ColPos_Note            = 11;
    this.ColPos_Occupied        = 12;
    this.ColPos_ValidContract   = 13;

    this.itemPack               = item;
    this.itemPackMaxLen         = 13;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.itemNo       = item[0];
      this.rentProperty = item[1];
      this.tenantName   = item[2];
      this.rentArrear   = item[3];
      this.deposit      = item[4];
      this.curRent      = item[5];
      this.dayRest      = item[6];
      this.contractNo   = item[7];
      this.accBalance   = item[8];
      this.status       = item[9];
      this.note         = item[10];
      this.occupied     = item[11];
      this.validContract= item[12];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemRptEvent] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.itemNo       = upd[0];
    this.rentProperty = upd[1];
    this.tenantName   = upd[2];
    this.rentArrear   = upd[3];
    this.deposit      = upd[4];
    this.curRent      = upd[5];
    this.dayRest      = upd[6];
    this.contractNo   = upd[7];
    this.accBalance   = upd[8];
    this.status       = upd[9];
    this.note         = upd[10];
    this.occupied     = upd[11];
    this.validContract= upd[12];
    
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  show(){
    var text = `itemRptStatus: \n(itemNo=${this.itemNo}, rentProperty=${this.rentProperty}, occupied=${this.occupied},tenantName=${this.tenantName},rentArrear=${this.rentArrear},deposit=${this.deposit},curRent=${this.curRent},dayRest=${this.dayRest},validContract=${this.validContract},contractNo=${this.contractNo},accBalance=${this.accBalance},status=${this.status},note=${this.note})`;
    return text;
  };

}

class itemRptEvent {
  constructor (item){

    this.itemNo          = null;
    this.date            = null;
    this.rentProperty    = null;
    this.tenantName      = null;
    this.contractNo      = null;
    this.event           = null;
    this.amount          = null;

    this.ColPos_Date            = 2;
    this.ColPos_RentProperty    = 3;
    this.ColPos_TenantName      = 4;
    this.ColPos_ContractNo      = 5;
    this.ColPos_Event           = 6;
    this.ColPos_Amount          = 7;

    this.itemPack               = item;
    this.itemPackMaxLen         = 7;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.itemNo          = item[0];
      this.date            = item[1];
      this.rentProperty    = item[2];
      this.tenantName      = item[3];
      this.contractNo      = item[4];
      this.event           = item[5];
      this.amount          = item[6];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemRptEvent] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }

  };

  update (upd) {
    this.itemNo          = upd[0];
    this.date            = upd[1];
    this.rentProperty    = upd[2];
    this.tenantName      = upd[3];
    this.contractNo      = upd[4];
    this.event           = upd[5];
    this.amount          = upd[6];
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  show(){
    var text = `itemRptEvent: \n(itemNo=${this.itemNo}, date=${this.date}, rentProperty=${this.rentProperty},tenantName=${this.tenantName},contractNo=${this.contractNo},event=${this.event},amount=${this.amount})`;
    return text;
  };

}