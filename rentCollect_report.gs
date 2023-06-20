/////////////////////////////////////////
// Golbal Var
/////////////////////////////////////////
const ReportEventValidContract = true; // only show the valid contract

const Color_Grey          = "#DCDCDC";
const Color_Yellow        = "#FFFF99";
const Color_Red           = "#F08080";
const Color_Green         = "#90EE90";

class itemRptStatus {
  constructor (item) {
    this.itemNo;
    this.rentProperty;
    this.occupied;
    this.tenantName;
    this.rentArrear;
    this.deposit;
    this.curRent;
    this.dayRest;
    this.validContract;
    this.contractNo;
    this.accBalance;
    this.status;

    this.ColPos_RentProperty    = 2;
    this.ColPos_Occupied        = 3;
    this.ColPos_ValidContract   = 4;
    this.ColPos_TenantName      = 5;
    this.ColPos_RentArrear      = 6;
    this.ColPos_Deposit         = 7;
    this.ColPos_CurRent         = 8;
    this.ColPos_DayRest         = 9;
    this.ColPos_ContractNo      = 10;
    this.ColPos_AccBalance      = 11;
    this.ColPos_Status          = 12;

    this.itemPack               = item;
    this.itemPackMaxLen         = 12;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.itemNo       = item[0];
      this.rentProperty = item[1];
      this.occupied     = item[2];
      this.validContract= item[3];
      this.tenantName   = item[4];
      this.rentArrear   = item[5];
      this.deposit      = item[6];
      this.curRent      = item[7];
      this.dayRest      = item[8];
      this.contractNo   = item[9];
      this.accBalance   = item[10];
      this.status       = item[11];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemRptEvent] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.itemNo       = upd[0];
    this.rentProperty = upd[1];
    this.occupied     = upd[2];
    this.validContract= upd[3];
    this.tenantName   = upd[4];
    this.rentArrear   = upd[5];
    this.deposit      = upd[6];
    this.curRent      = upd[7];
    this.dayRest      = upd[8];
    this.contractNo   = upd[9];
    this.accBalance   = upd[10];
    this.status       = upd[11];
    
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  show(){
    var text = `itemRptStatus: \n(itemNo=${this.itemNo}, rentProperty=${this.rentProperty}, occupied=${this.occupied},tenantName=${this.tenantName},rentArrear=${this.rentArrear},deposit=${this.deposit},curRent=${this.curRent},dayRest=${this.dayRest},validContract=${this.validContract},contractNo=${this.contractNo},accBalance=${this.accBalance},status=${this.status})`;
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

function rentCollect_report() {
  report_status();
  report_event();
}

function report_status() {
  /////////////////////////////////////////
  // README
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const topRowOfs = 1; // the offset from the top row, A2 is 1

  /////////////////////////////////////////
  // Update Status
  /////////////////////////////////////////
  if (SheetRptStatusName.getLastRow()>1) SheetRptStatusName.getRange(1+topRowOfs,1,SheetRptStatusName.getLastRow()-topRowOfs,SheetRptStatusName.getLastColumn()).clear(); // in case of empty
  for (var i=0;i<GLB_Property_arr.length;i++){
    var item = new itemRptStatus([]);
    var property = new itemProperty(GLB_Property_arr[i]);
    Logger.log(`property for status: ${property.show()}`);

    var itemNo = i;
    var rentProperty  = property.rentProperty;
    var occupied      = property.occupied;//SheetPropertyName.getRange(1+topRowOfs+i,property.ColPos_Occupied).getValue();
    var isValidContract = property.validContract;//SheetPropertyName.getRange(1+topRowOfs+i,property.ColPos_ValidContract).getValue();
    if (isValidContract) {
      var tenantName    = property.tenantName;// SheetPropertyName.getRange(1+topRowOfs+i,property.ColPos_TenantName).getValue();
      var contractNo    = property.contractNo;// SheetPropertyName.getRange(1+topRowOfs+i,property.ColPos_ContractNo).getValue();
      var contract      = new itemContract(GLB_Contract_arr[findContractNoPos(contractNo)]);
      var rentArrear    = contract.rentArear;// SheetContractName.getRange(1+topRowOfs+contractNo,contract.ColPos_RentArrear).getValue();
      var deposit       = contract.deposit;// SheetContractName.getRange(1+topRowOfs+contractNo,contract.ColPos_Deposit).getValue();
      var curRent       = property.curRent;// SheetPropertyName.getRange(1+topRowOfs+i,property.ColPos_CurRent).getValue();
      var dayRest       = property.dayRest;// SheetPropertyName.getRange(1+topRowOfs+i,property.ColPos_DayRest).getValue();

      // GLB_Contract_arr.forEach (
      //   function(data) {
      //     var item = new itemContract(data);
      //     Logger.log(`IUY: Property.itemNo: ${i}, contractNo: ${contractNo}, rentArrear: ${rentArrear}, item: ${item.show()}`);
      //   }
      // )
    } 
    else {
      var tenantName    = "N/A";
      var contractNo    = "N/A";
      var contract      = "N/A";
      var rentArrear    = "N/A";
      var deposit       = "N/A";
      var curRent       = "N/A";
      var dayRest       = "N/A";
    }

    // accumulate all bank record in this property
    var accBalance = 0;
    for (var j=0;j < GLB_BankRecord_arr.length;j++){
      var record = new itemBankRecord(GLB_BankRecord_arr[j]);
      if (record.rentProperty == property.rentProperty) accBalance += record.amount;
    }
    for (var j=0;j < GLB_MiscCost_arr.length;j++){
      var misc = new itemMiscCost(GLB_MiscCost_arr[j]);
      if (misc.rentProperty == property.rentProperty) {
        // if (misc.type == misc.MiscType_Charge_Fee)  accBalance += misc.amount;
        // if (misc.type == misc.MiscType_CashRent)    accBalance += misc.amount;
        // if (misc.type == misc.MiscType_Repare_Fee)  accBalance -= misc.amount;
        // if (misc.type == misc.MiscType_Refund)      accBalance -= misc.amount;
        accBalance += misc.balance_misc();
      }
    }
    
    
    
    SheetRptStatusName.getRange(1+topRowOfs+i,1).setValue(itemNo);
    SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_RentProperty).setValue(rentProperty);
    SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Occupied).setValue(occupied);
    SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(isValidContract);
    SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_AccBalance).setValue(accBalance);
    
    // Link the information from sheets
    if (occupied) { // if not occupied, left the blank along
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(tenantName);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_RentArrear).setValue(rentArrear);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Deposit).setValue(deposit);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_CurRent).setValue(curRent);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_DayRest).setValue(dayRest);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contractNo);
    }
    else if (isValidContract) { // may have end of cnstract but with unpaid rents case.
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(tenantName);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_RentArrear).setValue(rentArrear);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Deposit).setValue(deposit);
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contractNo);
    }
    
    // Logger.log(`contract: ${contract.show()}`);
    // report status with color
    if ((property.occupied == false) && (property.validContract == false)) {
      var status = "9.Vacancy";
      SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
      SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Grey);
    }
    else if ((property.occupied == false) && (property.validContract == true)) {
      if ((rentArrear) > 0) {
        var rptNum = (1000000 + rentArrear).toPrecision(7).toString().substring(1); // for add leading zero
        var status = `5.Need to refund by ${rptNum}.`;
        SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Yellow);
      }
      else if ((rentArrear) < 0) {
        var rptNum = (1000000 + Math.abs(rentArrear)).toPrecision(7).toString().substring(1); // for add leading zero
        var status = `8.Need to charge by ${rptNum}.`;
        SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Red);
      }
      else {
        var status = `7.Need to final check.`;
        SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Yellow);
      }
    }
    else if ((property.occupied == true) && (property.validContract == true)) {
      if (rentArrear + property.curRent*CFG_Val_obj["CFG_ReportArrearMargin"] < 0) {
        var rptNum = (1000000 + Math.abs(rentArrear)).toPrecision(7).toString().substring(1); // for add leading zero
        var status = `6.Rent arear is -${rptNum}, ${Math.floor(rentArrear*10/property.curRent)/10} month.`;
        SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Red);
      }
      else if (property.occupied && property.dayRest <= 30) {
        var status = `4.Contract near end within ${property.dayRest} days.`;
        SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Yellow);
      }
      else {
        var status = "0.Good!!!";
        SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Green);
      }
    }
    else {
      if (isValidContract == false) {var errMsg = `[rentCollect_report] How come a occupied property w/o valid contract, statusNo: ${item.itemNo}`;reportErrMsg(errMsg);}
    }
  
    // collect into array
    item.update([itemNo,rentProperty,occupied,isValidContract,tenantName,rentArrear,deposit,curRent,dayRest,contractNo,accBalance,status]);
    GLB_RptStatus_arr.push(item.itemPack);
  
  }

  // GLB_RptStatus_arr.forEach(
  //   function (data){
  //     item = new itemRptStatus(data);
  //     Logger.log(`RptStatus: ${item.show()}`);
  //   }
  // )

  // sort with priority
  var report = new itemRptStatus([]);
  var range = SheetRptStatusName.getRange(1+topRowOfs,1,SheetRptStatusName.getLastRow()-topRowOfs,SheetRptStatusName.getLastColumn());
  range.sort({column:report.ColPos_RentProperty,ascending: true});
  range.sort({column:report.ColPos_RentArrear,ascending: true});
  range.sort({column:report.ColPos_Status,ascending: false});
  
}

function report_event() {
  /////////////////////////////////////////
  // README
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const topRowOfs = 1; // the offset from the top row, A2 is 1

  /////////////////////////////////////////
  // Update Event
  /////////////////////////////////////////
  if (SheetRptEventName.getLastRow()>1) SheetRptEventName.getRange(1+topRowOfs,1,SheetRptEventName.getLastRow()-topRowOfs,SheetRptEventName.getLastColumn()).clear(); // in case of empty
  
  var itemNo          = 0;

  for (var i=0;i<GLB_Contract_arr.length;i++){
    var contract = new itemContract(GLB_Contract_arr[i]);
    Logger.log(`contract for event: ${contract.show()}`)
    if ((contract.validContract) && (contract.fromDate <= CONST_TODAY_DATE)) {
      // Event: start contract with deposit
      var item    = new itemRptEvent([]);
      var fromDate= Utilities.formatDate(contract.fromDate, "GMT+8", "yyyy/MM/dd");
      var toDate  = Utilities.formatDate(contract.toDate, "GMT+8", "yyyy/MM/dd");
      var event   = `1. Start of the contract from ${fromDate} to ${toDate} with deposit = ${contract.deposit}`;
      var amount  = -1 * contract.deposit;
      var upd     = [itemNo,fromDate,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
      item.update(upd);
      // SheetRptEventName.appendRow(item.itemPack);
      GLB_RptEvent_arr.push(item.itemPack);
      itemNo += 1;

      // Event: rent bill
      var finishDate;
      if (contract.toDate <= CONST_TODAY_DATE) finishDate = contract.toDate;
      else finishDate = new Date(CONST_TODAY_DATE.getTime()+CONST_MILLIS_PER_DAY);
      
      for (var j = new Date(contract.fromDate);j < finishDate;) {
        var item    = new itemRptEvent([]);
        var date    = Utilities.formatDate(j, "GMT+8", "yyyy/MM/dd");
        var event   = `2. Rent bill is ${contract.amount}`;
        var amount  = -1 * contract.amount;
        var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
        item.update(upd);
        // SheetRptEventName.appendRow(item.itemPack);
        GLB_RptEvent_arr.push(item.itemPack);
        itemNo += 1;

        j.setMonth(j.getMonth() + contract.period);

      }

      // Event: util bill
      for (var j=0;j<GLB_UtilBill_arr.length;j++) {
        var util = new itemUtilBill(GLB_UtilBill_arr[j]);
        if ((contract.fromDate <= util.date) && (util.date < contract.toDate) && (util.rentProperty == contract.rentProperty)){
          var item    = new itemRptEvent([]);
          var date    = Utilities.formatDate(util.date, "GMT+8", "yyyy/MM/dd");
          var event   = `3. Util bill is ${util.amount}`;
          var amount  = -1 * util.amount;
          var upd     = [itemNo,util.date,util.rentProperty,contract.tenantName,contract.itemNo,event,amount];
          item.update(upd);
          // SheetRptEventName.appendRow(item.itemPack);
          GLB_RptEvent_arr.push(item.itemPack);
          itemNo += 1;
        }
      }
      
      // Event: misc bill
      for (var j=0;j<GLB_MiscCost_arr.length;j++) {
        var misc = new itemMiscCost(GLB_MiscCost_arr[j]);
        if ((contract.fromDate <= misc.date) && (misc.date < contract.toDate) && (misc.rentProperty == contract.rentProperty)){
          var item    = new itemRptEvent([]);
          var date    = Utilities.formatDate(misc.date, "GMT+8", "yyyy/MM/dd");
          var event   = `4. Misc cost, type = ${misc.type}.`;
          var amount  = -1 * misc.expect_misc(); // due to the rentArrear is to sub this item
          var upd     = [itemNo,date,misc.rentProperty,contract.tenantName,contract.itemNo,event,amount];
          item.update(upd);
          // SheetRptEventName.appendRow(item.itemPack);
          GLB_RptEvent_arr.push(item.itemPack);
          itemNo += 1;
        }
      }
      
      // Event: bank record
      for (var j=0;j<GLB_BankRecord_arr.length;j++) {
        var record = new itemBankRecord(GLB_BankRecord_arr[j]);
        
        // for contractOverrid case
        var fromDate = new Date(contract.fromDate.getTime()-CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_FromDateMargin"]);
        var toDate   = new Date(finishDate.getTime()       +CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_ToDateMargin"]);
        if ((fromDate <= record.date) && (record.date < toDate)){
          if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == contract.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"")) { // for manually assign contract No record
            var item    = new itemRptEvent([]);
            var date    = Utilities.formatDate(record.date, "GMT+8", "yyyy/MM/dd");
            var event   = `5. Bank record.`;
            var amount  = record.amount;
            var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
            item.update(upd);
            // SheetRptEventName.appendRow(item.itemPack);
            GLB_RptEvent_arr.push(item.itemPack);
            itemNo += 1;
          }
        }

        // for normal case
        if ((contract.fromDate.getTime() <= record.date) && (record.date < finishDate)){
          if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
            if (contract.tenantAccount_arr.indexOf(record.fromAccount.toString())!=-1) {
              var item    = new itemRptEvent([]);
              var date    = Utilities.formatDate(record.date, "GMT+8", "yyyy/MM/dd");
              var event   = `5. Bank record.`;
              var amount  = record.amount;
              var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
              item.update(upd);
              // SheetRptEventName.appendRow(item.itemPack);
              GLB_RptEvent_arr.push(item.itemPack);
              itemNo += 1;
            }
          }
        }
      }

      // Event: The contract is over but not endContract
      if (contract.toDate <= CONST_TODAY_DATE) {
          var item    = new itemRptEvent([]);
          var date    = Utilities.formatDate(contract.toDate, "GMT+8", "yyyy/MM/dd");
          var event   = `6. Contract exceed last data ${date}, but not endContract!`;
          var amount  = null;
          var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
          item.update(upd);
          // SheetRptEventName.appendRow(item.itemPack);
          GLB_RptEvent_arr.push(item.itemPack);
          itemNo += 1;
      }

      // Event: The latest status
      for (var j=0;j<GLB_RptStatus_arr.length;j++) {
        var rptStatus = new itemRptStatus(GLB_RptStatus_arr[j]);
        if (rptStatus.rentProperty == contract.rentProperty) {
          var item    = new itemRptEvent([]);
          var date    = Utilities.formatDate(CONST_TODAY_DATE, "GMT+8", "yyyy/MM/dd");
          var event   = `9. Current stauts: ${rptStatus.status}`;
          var amount  = null;
          var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
          item.update(upd);
          // SheetRptEventName.appendRow(item.itemPack);
          GLB_RptEvent_arr.push(item.itemPack);
          itemNo += 1;
        }
      }
    
    }
  }
  
  // show on the sheet
  var item    = new itemRptEvent([]);
  var range = SheetRptEventName.getRange(1+topRowOfs,1,GLB_RptEvent_arr.length,item.itemPackMaxLen).setValues(GLB_RptEvent_arr);

  // sort by property and date for stable event order 
  var range = SheetRptEventName.getRange(1+topRowOfs,1,SheetRptEventName.getLastRow()-topRowOfs,SheetRptEventName.getLastColumn());
  range.sort({column:item.ColPos_Event,ascending: false});
  range.sort({column:item.ColPos_RentProperty,ascending: false});
  range.sort({column:item.ColPos_Date,ascending: false});
}

function findContractNoPos(contractNo) {
  for(var i=0;i<GLB_Contract_arr.length;i++){
    var item = new itemContract(GLB_Contract_arr[i]);
    if (item.itemNo == contractNo) return i;
  }

  if (1) {var errMsg = `[findContractNoPos] contractNo: ${contractNo} not found!?`; reportErrMsg(errMsg);}
  return -1
}
