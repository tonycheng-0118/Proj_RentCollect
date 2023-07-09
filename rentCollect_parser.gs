/////////////////////////////////////////
// Global Var
/////////////////////////////////////////
const CONST_CTBC_actWithdrawList = ["現金","中信卡","委代扣綜所稅"]; // A list to collect all of the known withdraw action apart from Tranfer
const CONST_CTBC_actDepositList  = ["現金", "利息", "委代入","現金存款機","委代入補助款","委代入貨物稅"]; // A list to collect all of the known deposit action apart from Tranfer
const CONST_KTB_actList = ["現金","現金D","現金E","本交A","利息D","退票D","稅款D"];
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

/////////////////////////////////////////
// Class
/////////////////////////////////////////
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
    this.contractNo             = null;
    this.tenantName             = null;
    this.validContract          = null;
    this.ColPos_ContractNo      = 6;
    this.ColPos_TenantName      = 7;
    this.ColPos_ValidContract   = 8;
    
    this.itemPack               = item;
    this.itemPackMaxLen         = 8;
        
    if (this.itemPack.length == this.itemPackMaxLen) {
      this.contractNo             = item[5];
      this.tenantName             = item[6];
      this.validContract          = item[7];
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
    var text = `itemUtilBill: \n(itemNo=${this.itemNo},date=${this.date},rentProperty=${this.rentProperty},amount=${this.amount},note=${this.note},contractNo=${this.contractNo},tenantName=${this.tenantName},validContract=${this.validContract})`;
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
    this.contractNo       = null;
    this.tenantName       = null;
    this.validContract    = null;
    
    this.ColPos_ContractNo      = 7;
    this.ColPos_TenantName      = 8;
    this.ColPos_ValidContract   = 9;

    this.MiscType_Charge_Fee    = `0.Charge_Fee`;
    this.MiscType_CashRent      = `1.Cash_Rent`;
    this.MiscType_SubRent       = `2.Sub_Rent`;
    this.MiscType_Commission    = `3.Commission`
    this.MiscType_Repare_Fee    = `5.Repare_Fee`;
    this.MiscType_Refund        = `6.Refund`;
    
    this.itemPack               = item;
    this.itemPackMaxLen         = 9;
    
    if (this.itemPack.length == this.itemPackMaxLen) {
      this.contractNo             = item[6];
      this.tenantName             = item[7];
      this.validContract          = item[8];
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
    var text = `itemMiscCost: \n(itemNo=${this.itemNo},date=${this.date},rentProperty=${this.rentProperty},amount=${this.amount},type=${this.type},note=${this.note},contractNo=${this.contractNo},tenantName=${this.tenantName},validContract=${this.validContract})`;
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
    this.contractOverrid        = item[11];
    this.contractNo             = null;
    this.rentProperty           = null;
    this.ColPos_RecordCheck     = 11;
    this.ColPos_ContractNo      = 13;
    this.ColPos_rentProperty    = 14;
    
    this.itemPack               = item;
    this.itemPackMaxLen         = 14;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.contractNo           = item[12];
      this.rentProperty         = item[13];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemBankRecord] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.contractNo             = upd[0];
    this.rentProperty           = upd[1];
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
    // Logger.log(`itemBankRecord update here @ ${this.itemNo}, upd.length: ${upd.length}, item.length: ${this.itemPack.length}`);
  };

  compare (item){
    var that = new itemBankRecord(item);
    // other.show();
    if (
      (this.date.toString()         == that.date.toString()) &&
      (this.action                  == that.action) && 
      (this.amount                  == that.amount) &&
      (this.fromAccountName         == that.fromAccountName) &&
      (this.fromAccount             == that.fromAccount) &&
      (this.toAccountName           == that.toAccountName) &&
      (this.toAccount               == that.toAccount) &&
      (this.balance                 == that.balance)
    ) {
      Logger.log(`AAA: ${this.show()}`);
      Logger.log(`BBB: ${that.show()}`);
      return true;
    } else return false
  }

  show(){
    var text = `itemBankRecord: \n(itemNo=${this.itemNo},date=${this.date},action=${this.action},amount=${this.amount},balance=${this.balance},fromAccountName=${this.fromAccountName},fromAccount=${this.fromAccount},toAccountName=${this.toAccountName},toAccount=${this.toAccount},mergeDate=${this.mergeDate},recordCheck=${this.recordCheck},contractOverrid=${this.contractOverrid},contractNo=${this.contractNo},rentProperty=${this.rentProperty})`;
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
    this.endDate                  = item[15];
    this.validContract            = null;
    this.rentArear                = null;
    this.dayRest                  = null;
    
    this.ColPos_ItemNo            = 1;
    this.ColPos_Deposit           = 5;
    this.ColPos_EndDate           = 15;
    this.ColPos_ValidContract     = 16;
    this.ColPos_RentArrear        = 17;
    this.ColPos_DayRest           = 18;

    this.itemPack                 = item;
    this.itemPackMaxLen           = 19;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.validContract          = item[16];
      this.rentArear              = item[17];
      this.dayRest                = item[18];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemContract] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.endDate                  = upd[0];
    this.validContract            = upd[1];
    this.rentArear                = upd[2];
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
    var text = `itemContract: \n(contractNo=${this.itemNo},fromDate=${this.fromDate},toDate=${this.toDate},rentProperty=${this.rentProperty},deposit=${this.deposit},amount=${this.amount},period=${this.period},tenantName=${this.tenantName},tenantAccountName_regex=${this.tenantAccountName_regex},tenantAccount_arr=${this.tenantAccount_arr},toAccountName=${this.toAccountName},toAccount=${this.toAccount},endContract=${this.endContract},note=${this.note},fileLink=${this.fileLink},endDate=${this.endDate},validContract=${this.validContract},rentArear=${this.rentArear},dayRest=${this.dayRest})`;
    // Logger.log(text);
    return text;
  };

}

/////////////////////////////////////////
// Function
/////////////////////////////////////////
function rentCollect_parser() {
  rentCollect_parser_Record();
  rentCollect_parser_Tenant();
  rentCollect_parser_UtilBill();
  rentCollect_parser_MiscCost();
  rentCollect_parser_Property();
  rentCollect_parser_Contract();
  rentCollect_parser_chkDuplicated();
}

function rentCollect_parser_Record() {
  rentCollect_parser_Record_ESUN(); // 玉山銀行
  rentCollect_parser_Record_KTB(); // 京城銀行
  rentCollect_parser_Record_CTBC(); // 中國信託
}

function rentCollect_parser_Record_ESUN() { // 玉山銀行
  /////////////////////////////////////////
  // README
  // With the upcoming new items in sheet.Import, this function will parse them and merge them with the sheet.Database.
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const importRowOfs = 9; // the offset from the top row, A1 is ofs 1
  const importRowSub = 3;  // the last row to ignore
  const importColOfs = 0;  // the offset from the left col, A1 is ofs 0
  const importContentLen = 9;
  const records_obj = [["玉山1001968210899","玉山0899","1001968210899"]];
  // for (const [k,v] of Object.entries(records_obj)) {
  //   Logger.log(`key: ${k}, value: ${v[1]}`);
  // }
  
  /////////////////////////////////////////
  // Import from source sheet
  /////////////////////////////////////////
  for (const [k,v] of Object.entries(records_obj)) {
    
    var isImportValid = rentCollect_import(v[0],false);

    if (isImportValid) {
      var data = SheetImportName.getRange(1+importRowOfs, 1+importColOfs, SheetImportName.getLastRow()-importRowOfs-importRowSub, importContentLen).getValues();
      for(i=0;i<data.length;i++){
        // itemNo
        var itemNo = i;

        // date
        if (typeof(data[i][0])==`string`) var date = new Date(data[i][0].toString().replace(/\*/g,"")); // to eliminate leading *
        else var date = new Date(data[i][0]);

        // action
        if ((data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")!="") & (data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")=="")) {
          // when 支出 non-empty and 存入 is empty, then it is withdraw
          withdraw = true;
          deposit  = false;
        } 
        else if ((data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")=="") & (data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")!="")) {
          // when 支出 empty and 存入 is non-empty, then it is deposit
          withdraw = false;
          deposit  = true;
        }
        else {
          Logger.log(`withdraw: ${withdraw}, deposit: ${deposit}, a:${data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")}, b:${data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")}`)
          if (1) {var errMsg = `[rentCollect_parser_Record_KTB] For ${v[0]} @ date: ${date}, no withdraw or deposit found!`; reportErrMsg(errMsg);}
        }
        var act = action_mapping("ESUN", i, data[i][2].toString(), withdraw, deposit);

        // amount
        var amount = (act == "TransferOut" || act == "Withdraw" || act == "OtherOut") ? data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"") : data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"");
        // Logger.log("i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);

        // balance
        var balance = data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"");
        
        // Logger.log("CCC: i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);
        // from
        var fromNote = note_mapping("ESUN", i, act, [data[i][2].toString(),data[i][6].toString(),data[i][7].toString(),data[i][8].toString()]);
        
        if (act != ""){
          GLB_Import_arr.push([date,act,amount,balance,fromNote[0],fromNote[1]]);
        }
        
        // Logger.log(`date:${date}, act:${act}, amount: ${amount}, balance: ${balance}, accountName: ${fromNote[0]}, account: ${fromNote[1]}`);
      }

      merge_BankRecord(v[1],v[2]);
    }
  }
  
}

function rentCollect_parser_Record_KTB() { // 京城銀行
  /////////////////////////////////////////
  // README
  // With the upcoming new items in sheet.Import, this function will parse them and merge them with the sheet.Database.
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const importRowOfs = 1; // the offset from the top row, A1 is ofs 0
  const importColOfs = 0; // the offset from the left col, A1 is ofs 0
  const importContentLen = 8;
  const records_obj = [["台幣交易明細查詢_9859","京城9859","005220099859"],["台幣交易明細查詢_9141","京城9141","005228009141"],["台幣交易明細查詢_8135","京城8135","005228008135"],["台幣交易明細查詢_3835","京城3835","005220113835"]];
  // for (const [k,v] of Object.entries(records_obj)) {
  //   Logger.log(`key: ${k}, value: ${v[1]}`);
  // }
  
  /////////////////////////////////////////
  // Import from source sheet
  /////////////////////////////////////////
  for (const [k,v] of Object.entries(records_obj)) {
    
    var isImportValid = rentCollect_import(v[0],false);

    if (isImportValid) {
      var data = SheetImportName.getRange(1+importRowOfs, 1+importColOfs, SheetImportName.getLastRow()-importRowOfs, importContentLen).getValues();
      for(i=0;i<data.length;i++){
        // itemNo
        var itemNo = i;

        // date
        var date = new Date(data[i][0]);

        // action
        if ((data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")!="") & (data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")=="")) {
          // when 支出 non-empty and 存入 is empty, then it is withdraw
          withdraw = true;
          deposit  = false;
        } 
        else if ((data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")!="") & (data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")!="")) {
          // when 支出/存入 are not empty, then it is deposit
          withdraw = false;
          deposit  = true;
        }
        else if ((data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")=="") & (data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")!="")) {
          // when 支出 empty and 存入 is non-empty, then it is deposit
          withdraw = false;
          deposit  = true;
        }
        else {
          Logger.log(`withdraw: ${withdraw}, deposit: ${deposit}, a:${data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")}, b:${data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")}`)
          if (1) {var errMsg = `[rentCollect_parser_Record_KTB] For ${v[0]} @ date: ${date}, no withdraw or deposit found!`; reportErrMsg(errMsg);}
        }
        var act = action_mapping("KTB", i, data[i][2].toString(), withdraw, deposit);

        // amount
        var amount = (act == "TransferOut" || act == "Withdraw") ? data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"") : data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"");
        // Logger.log("i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);

        // balance
        var balance = data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"");
        
        // Logger.log("CCC: i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);
        // from
        var fromNote = note_mapping("KTB", i, act, [data[i][3].toString(),data[i][7].toString()]);
        
        if (act != ""){
          GLB_Import_arr.push([date,act,amount,balance,fromNote[0],fromNote[1]]);
        }
        
        // Logger.log(`date:${date}, act:${act}, amount: ${amount}, balance: ${balance}, accountName: ${fromNote[0]}, account: ${fromNote[1]}`);
      }

      merge_BankRecord(v[1],v[2]);
    }
  }
  
}

function rentCollect_parser_Record_CTBC() { // 中國信託
  /////////////////////////////////////////
  // README
  // With the upcoming new items in sheet.Import, this function will parse them and merge them with the sheet.Database.
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  // const SheetImportName = SheetHandle.getSheetByName('Import');
  // const SheetDatabaseName = SheetHandle.getSheetByName('Database');
  
  /////////////////////////////////////////
  // Import from source sheet
  /////////////////////////////////////////
  rentCollect_import("DEPOSIT_APPLY_RECORD",false);

  /////////////////////////////////////////
  // Parse to itemRecord
  /////////////////////////////////////////
  // var GLB_Record_arr = new Array();
  var data = SheetImportName.getDataRange().getValues();
  for(i=0;i<data.length;i++){
    // itemNo
    var itemNo = i;

    // date
    var date = new Date(data[i][1]);

    // action
    var act = action_mapping("CTBC", i, data[i][2].toString(), data[i][3].toString(), data[i][4].toString());

    // amount
    var amount = (act == "TransferOut" || act == "Withdraw") ? data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"") : data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"");
    // Logger.log("i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);

    // balance
    var balance = data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    // Logger.log("CCC: i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);
    // from
    var fromNote = note_mapping("CTBC", i, act, data[i][6].toString());
    
    if (act != ""){
      GLB_Import_arr.push([date,act,amount,balance,fromNote[0],fromNote[1]]);
    }
  }

  merge_BankRecord("中國信託","000014853**1373*");
}

function merge_BankRecord(toAccountName,toAccount) {
  
  const bankRecordRowOfs = 1; // the offset from the top row, A2 is 1
  const bankRecordColOfs = 1; // to exclude itemNo
  const bankRecordContentLen = 11;
  
  /////////////////////////////////////////
  // Merge existed item
  /////////////////////////////////////////
  GLB_BankRecord_arr = new Array();  // Since GLB_BankRecord_arr is a Global array, will regen this content every time merge starting.
  if (SheetBankRecordName.getLastRow()) { // for a non-empty database case
    var data = SheetBankRecordName.getRange(1+bankRecordRowOfs, 1+bankRecordColOfs, SheetBankRecordName.getLastRow()-bankRecordRowOfs, bankRecordContentLen  ).getValues(); // exclude itemNo column
    for(i=0;i<data.length;i++){
      GLB_BankRecord_arr.push(data[i]);
    }

    var Compare_BankRecord_arr  = new Array();
    var compare_data = SheetBankRecordName.getRange(1+bankRecordRowOfs, 1+bankRecordColOfs, SheetBankRecordName.getLastRow()-bankRecordRowOfs, bankRecordContentLen-5).getValues(); // exclude itemNo column and [toAccountName, toAccount, mergeDate, RecordCheck, ContractOverrid]
    for(i=0;i<compare_data.length;i++){
      Compare_BankRecord_arr.push(compare_data[i]);
    }
  }
  
  for(i=0;i<GLB_Import_arr.length;i++){
    if (Compare_BankRecord_arr.join().indexOf(GLB_Import_arr[i].join()) == -1){
      GLB_BankRecord_arr.push(GLB_Import_arr[i].concat([toAccountName,toAccount,CONST_TODAY_DATE,"",""]));// expand a placeholder for RecordCheck, ContractOverrid
    }
    else {
      // Found duplicated itemRecord, passed.
    }
  }

  /////////////////////////////////////////
  // Sort itemRecord with Date in acending order
  /////////////////////////////////////////
  GLB_BankRecord_arr.sort(
    function(x,y){
      var xp = x[0];
      var yp = y[0];
      return xp == yp ? 0 : xp < yp ? -1 : 1;
    }
  )
  
  /////////////////////////////////////////
  // Attach with serial number
  /////////////////////////////////////////
  for(var i=0;i<GLB_BankRecord_arr.length;i++){
    GLB_BankRecord_arr[i] = [i].concat(GLB_BankRecord_arr[i]);
  }

  // GLB_BankRecord_arr.forEach (
  //   function(data) {
  //     var item = new itemBankRecord(data);
  //     Logger.log(`MNB: ${item.itemPack.length}, ${item.show()}`);
  //   }
  // )

  /////////////////////////////////////////
  // Write Out
  /////////////////////////////////////////
  SheetBankRecordName.getRange(1+bankRecordRowOfs,1,GLB_BankRecord_arr.length,GLB_BankRecord_arr[0].length).clearContent();
  SheetBankRecordName.getRange(1+bankRecordRowOfs,1,GLB_BankRecord_arr.length,GLB_BankRecord_arr[0].length).setValues(GLB_BankRecord_arr);
  Logger.log(`[Info] Size of GLB_BankRecord_arr is ${GLB_BankRecord_arr.length}`);
  
}

function note_mapping(type, id, act, note_in) {
  // Only handle TransferIn action
  
  if (type == "CTBC") {
    if (act == "TransferIn"){
      // rid of self account number
      var note_tmp = note_in.replace(CONST_This_Account_Number,"");

      // retrieve fromAccountName
      var regExp = new RegExp("([0-9]{9}\\*\\*[0-9]{4}\\*)","gi"); // escape word
      var fromAccount = regExp.exec(note_tmp)[0];
      var fromAccountName = note_tmp.replace(regExp,"").replace(/[\s|\n|\r|\t]/g,"").split("\n");
      
      return [fromAccountName[0],fromAccount];
    }
  }
  else if (type == "KTB") {
    if (act == "TransferIn"){
      // retrieve fromAccountName
      var regExp = new RegExp("([0-9]{10}\d*)","gi"); // escape word
      if (note_in[0].match(regExp) != null) {
        var fromAccount = (regExp.exec(note_in[0])[0]);
      }
      else {
        var fromAccount = "";
      }
      var fromAccountName = note_in[0].replace(/[\s|\n|\r|\t]/g,"") + ";" + note_in[1].replace(/[\s|\n|\r|\t]/g,"");
      
      return [fromAccountName,fromAccount];
    }
    else if (act == "Deposit"){
      // retrieve fromAccountName
      var regExp = new RegExp("([0-9]{10}\d*)","gi"); // escape word
      var fromAccount = "";//note_in[0].replace(/[\s|\n|\r|\t]/g,""); // for KTB, only AccountName is available
      var fromAccountName = note_in[0].replace(/[\s|\n|\r|\t]/g,"") + ";" + note_in[1].replace(/[\s|\n|\r|\t]/g,"");
      
      return [fromAccountName,fromAccount];
    }
  }
  else if (type == "ESUN"){
    if (act == "TransferIn"){
      // retrieve fromAccountName
      var regExp = new RegExp("([0-9]{3}/[0-9]{10}\d*)","gi"); // escape word
      if (note_in[2].match(regExp) != null) {
        var fromAccount = (regExp.exec(note_in[2])[0]);
      }
      else {
        var fromAccount = "";
      }
      var fromAccountName = note_in[1].replace(/[\s|\n|\r|\t]/g,"") + ";" + note_in[3].replace(/[\s|\n|\r|\t]/g,"");
      
      return [fromAccountName,fromAccount];
    }
    if ((act == "OtherIn") || (act == "OtherOut")){
      var fromAccount = "";
      var fromAccountName = note_in[0].replace(/[\s|\n|\r|\t]/g,""); // to prevent duplicate no used record
      return [fromAccountName,fromAccount];
    }
  }
  else {
    if (1) {var errMsg = `[note_mapping] import format does not support at id: ${id}?? act is: ${act}, note_in: ${note_in}`; reportErrMsg(errMsg);}
  }


  return ["",""]; // empty string

}

function action_mapping(type, id, act_in, withdraw, deposit){
  // return TransferOut, TransferIn, Withdraw, Deposit type only
  
  var act = act_in.replace(/[\s|\n|\r|\t]/g,"").split("\n");
  var isWithdraw = withdraw!="";
  var isDeposit  = deposit!="";

  // integrity check
  if ( isWithdraw &&  isDeposit) console.error("Withdraw and Deposit happened at the same time!!?");
  if (!isWithdraw && !isDeposit) console.error("Withdraw and Deposit missed!!?");
  
  if (type=="CTBC") {
    // looking for keyword "匯" or "轉"
    var regExp = new RegExp(".*[匯|轉].*","gi");
    var match = regExp.exec(act[0]);
    if (match != null) {
      if (isWithdraw) return "TransferOut";
      if (isDeposit)  return "TransferIn";
    }

    if (isWithdraw && CONST_CTBC_actWithdrawList.indexOf(act[0])!=-1){
      return "Withdraw"
      // Logger.log("HAHA in actWithdrawList: %d, act: %s",actWithdrawList.indexOf(act[0]), act[0]);
    }

    if (isDeposit && CONST_CTBC_actDepositList.indexOf(act[0])!=-1){
      return "Deposit"
      // Logger.log("HAHA in actWithdrawList: %d, act: %s",actWithdrawList.indexOf(act[0]), act[0]);
    }

    // report error if no condition met
    if (1) {var errMsg = `[action_mapping] Something untracked at id: ${id}?? act_in is: ${act_in}, act is: ${act}`; reportErrMsg(errMsg);}
  }
  else if (type=="KTB"){
    var regExp = new RegExp(".*[轉帳].*","gi");
    var match = regExp.exec(act[0]);
    if (match != null) {
      if (isWithdraw) return "TransferOut";
      if (isDeposit)  return "TransferIn";
    }

    if (isWithdraw && CONST_KTB_actList.indexOf(act[0])!=-1){
      return "Withdraw"
      // Logger.log("HAHA in actWithdrawList: %d, act: %s",actWithdrawList.indexOf(act[0]), act[0]);
    }

    if (isDeposit && CONST_KTB_actList.indexOf(act[0])!=-1){
      return "Deposit"
      // Logger.log("HAHA in actWithdrawList: %d, act: %s",actWithdrawList.indexOf(act[0]), act[0]);
    }

    if (1) {var errMsg = `[action_mapping] Something untracked at id: ${id}?? act_in is: ${act_in}, act is: ${act}`; reportErrMsg(errMsg);}
    return ""
  }
  else if (type=="ESUN"){
    var regExp = new RegExp(".*[匯|轉].*","gi");
    if (regExp.exec(act[0])!=null) {
      if (isWithdraw) return "TransferOut";
      if (isDeposit)  return "TransferIn";
    }

    var regExp = new RegExp(".*[存].*","gi");
    if (regExp.exec(act[0])!=null) {
      if (isWithdraw) return "Withdraw";
      if (isDeposit ) return "Deposit";
    }
    else { // like interest, load, bank adjustment, which is not record of interesting
      if (isWithdraw) return "OtherOut";
      if (isDeposit ) return "OtherIn";
    }

    if (1) {var errMsg = `[action_mapping] Something untracked at id: ${id}?? act_in is: ${act_in}, act is: ${act}`; reportErrMsg(errMsg);}
    return ""
  }
  else {
    if (1) {var errMsg = `[action_mapping]import format does not support at id: ${id}?? act_in is: ${act_in}, act is: ${act}`; reportErrMsg(errMsg);}
    return ""
  }

}

function rentCollect_parser_chkDuplicated() {
  for (i=0;i<GLB_BankRecord_arr.length-1;i++){ // less likely duplicated, check for assurance
    for (j=i+1;j<GLB_BankRecord_arr.length;j++){
      var a = new itemBankRecord(GLB_BankRecord_arr[i]);
      if (a.compare(GLB_BankRecord_arr[j])) {var errMsg = `[rentCollect_parser_chkDuplicated] BankRecord:[${i}]: ${GLB_BankRecord_arr[i]}, BankRecord:[${j}]: ${GLB_BankRecord_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

  for (i=0;i<GLB_Contract_arr.length-1;i++){
    for (j=i+1;j<GLB_Contract_arr.length;j++){
      var a = new itemContract(GLB_Contract_arr[i]);
      if (a.compare(GLB_Contract_arr[j])) {var errMsg = `[rentCollect_parser_chkDuplicated] Contract: ${GLB_Contract_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

  for (i=0;i<GLB_UtilBill_arr.length-1;i++){
    for (j=i+1;j<GLB_UtilBill_arr.length;j++){
      var a = new itemUtilBill(GLB_UtilBill_arr[i]);
      if (a.compare(GLB_UtilBill_arr[j])) {var errMsg = `[rentCollect_parser_chkDuplicated] UtilBill: ${GLB_UtilBill_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

  for (i=0;i<GLB_MiscCost_arr.length-1;i++){
    for (j=i+1;j<GLB_MiscCost_arr.length;j++){
      var a = new itemMiscCost(GLB_MiscCost_arr[i]);
      if (a.compare(GLB_MiscCost_arr[j])) {var errMsg = `[rentCollect_parser_chkDuplicated] MisoCost: ${GLB_MiscCost_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

  for (i=0;i<GLB_Property_arr.length-1;i++){
    for (j=i+1;j<GLB_Property_arr.length;j++){
      var a = new itemProperty(GLB_Property_arr[i]);
      if (a.compare(GLB_Property_arr[j])) {var errMsg = `[rentCollect_parser_chkDuplicated] Property: ${GLB_Property_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

}

function rentCollect_parser_Tenant() {
  /////////////////////////////////////////
  // Parse to itemTenant 
  /////////////////////////////////////////
  const topRowOfs = 1;
  const tenantNameColOfs = 1;
  const accountNameColOfs = 3;
  const accountColOfs = 4;
  var all_tenant_arr  = new Array();
  var all_account_arr = new Array();
  SheetTenantName.getRange(1+topRowOfs,1,SheetTenantName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  var data = SheetTenantName.getRange(1+topRowOfs,1,SheetTenantName.getLastRow()-topRowOfs,SheetTenantName.getLastColumn()).getValues();
  for(i=0;i<data.length;i++){
    var itemNo = i;
    
    var account_arr = new Array();
    if (data[i].length<accountColOfs+1) {var errMsg = `[rentCollect_parser_Tenant] The tenant account data is missing for ${itemNo}`; reportErrMsg(errMsg);}
    chkNotEmptyEntry(data[i][1]);
    for (var a=accountColOfs;a<data[i].length;a++){
      if (data[i][a].toString().replace(/ /g,'')!='') {
        var account = data[i][a].toString().replace(/ /g,'');
        account_arr.push(account);
        all_account_arr.push(account);
        chkValidAccount(account);
      }
    }

    var tenant = data[i][tenantNameColOfs].toString().replace(/[\s|\n|\r|\t]/g,"");
    var accountName = data[i][accountNameColOfs].toString().replace(/[\s|\n|\r|\t]/g,"");
    all_tenant_arr.push(tenant);
    GLB_Tenant_obj[tenant] = [accountName,account_arr]; // hash table, key=telantName, value=[Account]
  
    // Write TenantNo
    SheetTenantName.getRange(1+topRowOfs+i,1).setValue(itemNo);

  }

  
  /////////////////////////////////////////
  // chkValidTenant
  /////////////////////////////////////////
  // Logger.log(Object.keys(GLB_Tenant_obj).length);
  // Logger.log(Object.values(GLB_Tenant_obj).length);
  // Logger.log(Object.keys(itemTenant_obj).indexOf("Tony5"));
  // chkValidTenant(GLB_Tenant_obj);

  // for (const [k,v] of Object.entries(GLB_Tenant_obj)) {
  //   Logger.log(`key: ${k}, value[0]: ${v[0]}, value[1]: ${v[1]}`);
  //   Logger.log(`GLB_Tenant_obj[tenantName][0]: ${GLB_Tenant_obj[k][0]}`)
  //   Logger.log(`GLB_Tenant_obj[tenantName][1]: ${GLB_Tenant_obj[k][1]}`);
  // }
  
  // check duplicated tenant
  for (i=0;i<all_tenant_arr.length-1;i++){
    for (j=i+1;j<all_tenant_arr.length;j++){
      if (all_tenant_arr[i] == all_tenant_arr[j]) {var errMsg = `[rentCollect_parser_Tenant] tenant: ${all_tenant_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

  // check duplicated account
  for (i=0;i<all_account_arr.length-1;i++){
    for (j=i+1;j<all_account_arr.length;j++){
      if (all_account_arr[i] == all_account_arr[j]) {var errMsg = `[rentCollect_parser_Tenant] account: ${all_account_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }


}

function rentCollect_parser_Contract() {
  /////////////////////////////////////////
  // Parse to itemContract
  /////////////////////////////////////////
  var topRowOfs = 1;
  // SheetContractName.getRange(1+topRowOfs,1,SheetContractName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  var data = SheetContractName.getRange(1+topRowOfs,1,SheetContractName.getLastRow()-topRowOfs,SheetContractName.getLastColumn()).getValues();
  var getMaxContractNo = -1;
  for(i=0;i<data.length;i++){
    // itemNo
    var itemNo = data[i][0];

    // fromDate
    chkNotEmptyEntry(data[i][1]);
    var fromDate = new Date(data[i][1]);
    fromDate.setSeconds(0);fromDate.setMinutes(0);fromDate.setHours(0);

    // toDate
    var toDate = new Date(data[i][2]);
    toDate.setSeconds(0);toDate.setMinutes(0);toDate.setHours(0);
    chkNotEmptyEntry(data[i][2]);
    if (toDate <= fromDate) {var errMsg = `[rentCollect_contract] This is not a valid toDate: contractNo: ${itemNo}`;reportErrMsg(errMsg);}

    // rentProperty
    chkNotEmptyEntry(data[i][3]);
    var rentProperty = data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"");

    // deposit
    chkNotEmptyEntry(data[i][4]);
    var deposit = data[i][4];

    // amount
    chkNotEmptyEntry(data[i][5]);
    var amount = data[i][5];

    // period
    chkNotEmptyEntry(data[i][6]);
    var period = data[i][6];
    if (period==0) {var errMsg = `[rentCollect_contract] Period cannot be 0.`};

    // tenantName
    chkNotEmptyEntry(data[i][7]);
    var tenantName = data[i][7].toString().replace(/[\s|\n|\r|\t]/g,"");

    // tenantAccount_arr
    if (chkValidTenantName(tenantName)){
      var tenantAccountName_regex = GLB_Tenant_obj[tenantName][GLB_Tenant_AccountName_Pos];
      var tenantAccount_arr = GLB_Tenant_obj[tenantName][GLB_Tenant_Account_Pos];
    } 
    else {
      var tenantAccountName_regex = "";
      var tenantAccount_arr = []; // to not handing
    }

    // toAccountName
    chkNotEmptyEntry(data[i][8])
    var toAccountName = data[i][8].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    // toAccount
    chkValidAccount(data[i][9])
    var toAccount = data[i][9].toString().replace(/[\s|\n|\r|\t]/g,"");

    // endContract
    var tmp = data[i][10].toString().replace(/[\s|\n|\r|\t]/g,"");
    var endContract;
    if (tmp == "") endContract = false
    else if (tmp == "EndContract") endContract = true
    else {var errMsg = `[rentCollect_parser_Tenant] This is not a valid endContract: ${tmp}`; reportErrMsg(errMsg);}

    // note
    var note = data[i][11].toString().replace(/[\s|\n|\r|\t]/g,"");

    // fileLink
    var fileLink = data[i][12].toString().replace(/[\s|\n|\r|\t]/g,"");

    // endDate
    var chk_illegal = true;
    var endDate = data[i][13];
    
    if ((typeof(endDate) == `string`) && (endDate.replace(/ /g,'') == '')) chk_illegal = false;
    else if ((typeof(endDate) == `object`)) chk_illegal = false;
    else {var errMsg = `[rentCollect_parser_Tenant] This is not a valid endDate: ${endDate}`; reportErrMsg(errMsg);}
    
    GLB_Contract_arr.push([itemNo,fromDate,toDate,rentProperty,deposit,amount,period,tenantName,tenantAccountName_regex,tenantAccount_arr,toAccountName,toAccount,endContract,note,fileLink,endDate]);

    // var item = new itemContract(GLB_Contract_arr[i]);
    // Logger.log(`RRR: ${item.show()}`);

    // Write contractNo
    // SheetContractName.getRange(1+topRowOfs+i,1).setValue(itemNo);
    if (getMaxContractNo < itemNo) getMaxContractNo = itemNo; // for new contractNo assignment.

  }

  /////////////////////////////////////////
  // Post GLB_Contract_arr process
  /////////////////////////////////////////
  for(i=0;i<GLB_Contract_arr.length;i++){
    // to assign contractNo to new contract
    var item = new itemContract(GLB_Contract_arr[i]);
    if (item.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
      item.itemNo = ++getMaxContractNo;
      GLB_Contract_arr[i][item.ColPos_ItemNo-1] = item.itemNo;
      SheetContractName.getRange(1+topRowOfs+i,item.ColPos_ItemNo).setValue(item.itemNo);
    }

    // check and update end contract date
    if (item.endContract) { // to update the endDate
      if (item.endDate.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
        if (CONST_TODAY_DATE < item.fromDate) {
          if (1) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} EndContract incorrect`; reportErrMsg(errMsg);}
        }
        else if (CONST_TODAY_DATE < item.toDate) {
          var endDate = new Date(CONST_TODAY_DATE.getTime());
          GLB_Contract_arr[i][item.ColPos_EndDate] = endDate;
          SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(endDate); // abnormal end date, terminate contract ahead.
        }
        else if (item.toDate <= CONST_TODAY_DATE) {
          var endDate = new Date(item.toDate.getTime());
          GLB_Contract_arr[i][item.ColPos_EndDate] = endDate;
          SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(endDate); // normal end date
        }
        else {
          if (1) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} how come???`; reportErrMsg(errMsg);}
        }
      }
      else {
        // leave it what it is
      }
    }
    else {//if (item.endContract == false) { // should be empty
      GLB_Contract_arr[i][item.ColPos_EndDate] = CONST_SuperFeatureDate; // expected to be not used, just for data type consistency
      SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(""); 
    }
  
  
  }
  // GLB_Contract_arr.forEach (
  //   function(data) {
  //     var item = new itemContract(data);
  //     Logger.log(`RXX: ${item.show()}, getMaxContractNo: ${getMaxContractNo}`);
  //   }
  // )


  // integrity check cross contract and tenant
  chkContractIntegrity();
}

function chkValidTenantName(tenantName){

  var tenantName_arr = Object.keys(GLB_Tenant_obj);
  var match = false;
  for (var i =0;i<tenantName_arr.length;i++){
    if ((match == false) && (tenantName_arr[i] == tenantName)) match = true;
  }
  
  if (match==false) {var errMsg = `[chkValidTenantName] This is not a valid TenantName: ${tenantName}`;reportErrMsg(errMsg);}
  
  return match;
}

function chkValidAccount(account){
    var regExpCTBC = new RegExp("([0-9]{9}\\*\\*[0-9]{4}\\*)","gi"); // escape word
    var meetCTBC = regExpCTBC.exec(account);

    var regExpKTB = new RegExp("([0-9]{10}\d*)","gi");
    var meetCKTB = regExpKTB.exec(account);

    if (meetCTBC || meetCKTB) {
      return true;
    } 
    else {
      if (1) {var errMsg = `[chkValidAccount] This is not a valid account: ${account}`; reportErrMsg(errMsg);}
      return false;   
    }
}

function chkContractIntegrity(){
  // var item = new itemContract(itemIn);
  // Logger.log("AAA: %s",item.contractNo);
  // Logger.log(Object.keys(GLB_Tenant_obj).length);
  // 3. ToDate <= FromDate
  
  //Only one property can be rented in a given window.
  for(var i=0;i<GLB_Contract_arr.length-1;i++){
    for (var j=i+1;j<GLB_Contract_arr.length;j++){
      var a = new itemContract(GLB_Contract_arr[i]);
      var b = new itemContract(GLB_Contract_arr[j]);
      var a_preperty_date_valid = (a.fromDate <= CONST_TODAY_DATE) && (CONST_TODAY_DATE < a.toDate) && (a.endContract==false);
      var b_preperty_date_valid = (b.fromDate <= CONST_TODAY_DATE) && (CONST_TODAY_DATE < b.toDate) && (b.endContract==false);
      var a_preperty_valid = (a.endContract==false);
      var b_preperty_valid = (b.endContract==false);
      // Logger.log(`AAA: ${a.show()}`);
      // Logger.log(`BBB: ${b.show()}`);

      if ((a.rentProperty == b.rentProperty) && a_preperty_date_valid && b_preperty_date_valid){
        var overlap = ((a.fromDate <= b.toDate  ) && (b.toDate   < a.toDate  )) ||
                      ((a.fromDate <= b.fromDate) && (b.fromDate < a.toDate)) ||
                      ((b.fromDate <= a.toDate  ) && (a.toDate   < b.toDate  )) ||
                      ((b.fromDate <= a.fromDate) && (a.fromDate < b.toDate));
        if (overlap) {var errMsg = `[chkContractIntegrity] rentProperty date range overlap at contractNo: ${b.itemNo}. A:${a.show()}, B:${b.show()}`; reportErrMsg(errMsg);}
      }

      if ((a.rentProperty == b.rentProperty) && a_preperty_valid && b_preperty_valid) {
        if ((a_preperty_date_valid == false) && (b_preperty_date_valid == true)) {var warnMsg = `[chkContractIntegrity] This contractNo: ${a.itemNo} is end and duplicated to contractNo: ${b.itemNo}. A:${a.show()}, B:${b.show()}`; reportWarnMsg(warnMsg);} 
        // if (b_preperty_date_valid == false) {var errMsg = `[chkContractIntegrity] This contractNo: ${b.itemNo} is end and duplicated to contractNo: ${a.itemNo}. A:${a.show()}, B:${b.show()}`; reportErrMsg(errMsg);}
      }
    }
  }

  //RentProperty is existed
  for(var i=0;i<GLB_Contract_arr.length;i++){
    var match = false;
    var contract = new itemContract(GLB_Contract_arr[i]);
    for (var j=0;j<GLB_Property_arr.length;j++){
      var property = new itemProperty(GLB_Property_arr[j]);
      match = (contract.rentProperty == property.rentProperty);
      if (match) break;
    }
    // Logger.log(`AAA: ${contract.show()}`);
    // Logger.log(`match: ${match}`);
    if (match==false) {var errMsg = `[chkContractIntegrity] rentProperty no matched: ${contract.itemNo}. ${contract.show()}`; reportErrMsg(errMsg);}
  }

  //ContractOverrid correlate to existed contractNo
  for (i=0;i<GLB_BankRecord_arr.length;i++){ // less likely duplicated, check for assurance
    var record = new itemBankRecord(GLB_BankRecord_arr[i]);
    if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") != "") {
      if (findContractNoPos(record.contractOverrid)==-1) {var errMsg = `[chkContractIntegrity] ContractOverrid contractNo not existed: ${record.itemNo}. ${record.show()}`; reportErrMsg(errMsg);}
      else {
        var contract = new itemContract(GLB_Contract_arr[findContractNoPos(record.contractOverrid)]);

        var fromDate = new Date(contract.fromDate.getTime()-CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_FromDateMargin"]);
        var toDate   = new Date(contract.toDate.getTime()  +CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_ToDateMargin"]);
        
        if (!(fromDate<=record.date && record.date<toDate)) {var errMsg = `[chkContractIntegrity] ContractOverrid date not valid: ${record.itemNo}, should be within from ${fromDate} to ${toDate}. \n ${record.show()}. \n ${contract.show()}}`; reportErrMsg(errMsg);}
        
        if (record.toAccountName != contract.toAccountName) {var errMsg = `[chkContractIntegrity] ContractOverrid toAccountName diff: ${record.itemNo}. \n ${record.show()}. \n ${contract.show()}}`; reportErrMsg(errMsg);}

        if (record.toAccount != contract.toAccount) {var errMsg = `[chkContractIntegrity] ContractOverrid toAccount diff: ${record.itemNo}. \n ${record.show()}. \n ${contract.show()}}`; reportErrMsg(errMsg);}
      }
    }
  } 

}

function chkNotEmptyEntry(entry){
  var type = typeof(entry);
  if (type == `string`){
    if (entry.replace(/ /g,'')=='') {var errMsg = `[chkNotEmptyEntry] This is an empty entry: ${entry}`; reportErrMsg(errMsg);}
  }
  else if ((type == `object`) || (type == `number`) || (type == `boolean`)){
    if (entry == null) {var errMsg = `[chkNotEmptyEntry] This is an empty entry: ${entry}`; reportErrMsg(errMsg);}
  }
  else {
    if (1) {var errMsg = `[chkNotEmptyEntry] No typeof(${entry}): ${typeof(entry)})`; reportErrMsg(errMsg);}  
  }

  return true;
}

function rentCollect_parser_UtilBill() {
  /////////////////////////////////////////
  // Parse to itemUtilBill 
  /////////////////////////////////////////
  const topRowOfs = 1;
  SheetUtilBillName.getRange(1+topRowOfs,1,SheetUtilBillName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  var data = SheetUtilBillName.getRange(1+topRowOfs,1,SheetUtilBillName.getLastRow()-topRowOfs,SheetUtilBillName.getLastColumn()).getValues();
 
  for(i=0;i<data.length;i++){
    // itemNo
    var itemNo = i;
    
    // date
    chkNotEmptyEntry(data[i][1]);
    var date = new Date(data[i][1]);
    date.setSeconds(0);date.setMinutes(0);date.setHours(0);

    // rentProperty
    chkNotEmptyEntry(data[i][2]);
    var rentProperty = data[i][2];

    // amount
    chkNotEmptyEntry(data[i][3]);
    var amount = data[i][3];

    // note
    var note = data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    GLB_UtilBill_arr.push([itemNo,date,rentProperty,amount,note]);
    
    // var item = new itemUtilBill(GLB_UtilBill_arr[i]);
    // item.show()

    // Write UtilBillNo
    SheetUtilBillName.getRange(1+topRowOfs+i,1).setValue(itemNo);

  }

}

function rentCollect_parser_MiscCost() {
  /////////////////////////////////////////
  // Parse to itemMiscCost 
  /////////////////////////////////////////
  const topRowOfs = 1;
  SheetMiscCostName.getRange(1+topRowOfs,1,SheetMiscCostName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  var data = SheetMiscCostName.getRange(1+topRowOfs,1,SheetMiscCostName.getLastRow()-topRowOfs,SheetMiscCostName.getLastColumn()).getValues();
 
  for(i=0;i<data.length;i++){
    // itemNo
    var itemNo = i;
    
    // date
    chkNotEmptyEntry(data[i][1]);
    var date = new Date(data[i][1]);
    date.setSeconds(0);date.setMinutes(0);date.setHours(0);

    // rentProperty
    chkNotEmptyEntry(data[i][2]);
    var rentProperty = data[i][2];

    // amount
    chkNotEmptyEntry(data[i][3]);
    var amount = data[i][3];

    // type
    chkNotEmptyEntry(data[i][4]);
    var type = data[i][4];

    // note
    var note = data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    GLB_MiscCost_arr.push([itemNo,date,rentProperty,amount,type,note]);
    
    var item = new itemMiscCost(GLB_MiscCost_arr[i]);
    // item.show()

    // Write UtilBillNo
    SheetMiscCostName.getRange(1+topRowOfs+i,1).setValue(itemNo);

  }

}

function rentCollect_parser_Property() {
  /////////////////////////////////////////
  // Parse to itemProperty 
  /////////////////////////////////////////
  const topRowOfs = 1;
  SheetPropertyName.getRange(1+topRowOfs,1,SheetPropertyName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  var data = SheetPropertyName.getRange(1+topRowOfs,1,SheetPropertyName.getLastRow()-topRowOfs,SheetPropertyName.getLastColumn()).getValues();
 
  for(i=0;i<data.length;i++){
    // itemNo
    var itemNo = i;
    
    // rentProperty
    chkNotEmptyEntry(data[i][1]);
    var rentProperty = data[i][1];

    // note
    var note = data[i][2].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    GLB_Property_arr.push([itemNo,rentProperty,note]);
    
    // var item = new itemProperty(GLB_Property_arr[i]);
    // item.show()

    // Write UtilBillNo
    SheetPropertyName.getRange(1+topRowOfs+i,1).setValue(itemNo);

  }

}
