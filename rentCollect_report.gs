/////////////////////////////////////////
// Golbal Var
/////////////////////////////////////////
const ReportEventValidContract = true; // only show the valid contract

const Color_Grey          = "#DCDCDC";
const Color_Yellow        = "#FFFF99";
const Color_Red           = "#F08080";
const Color_Green         = "#90EE90";

class itemRptAnalysis {
  constructor (item) {
    this.itemNo;
    this.propertyGroup;
    this.monthAccRent;
    
    this.ColPos_ItemNo            = 1;
    this.ColPos_PropertyGroupName = 2;
    this.ColPos_PropertyGroup     = 3;
    this.ColPos_MonthAccRent      = 4;

    this.itemPack               = item;
    this.itemPackMaxLen         = 4;

    if (this.itemPack.length == this.itemPackMaxLen) {
      this.itemNo             = item[0];
      this.propertyGroupName  = item[1];
      this.propertyGroup      = item[2];
      this.monthAccRent       = item[3];
    }
    else if (this.itemPack.length > this.itemPackMaxLen) {
      if (1) {var errMsg = `[itemRptAnalysis] Too much itemPack.length: ${this.itemPack.length} @ itemNo: ${this.itemNo}`; reportErrMsg(errMsg);}
    }
  };

  update (upd) {
    this.itemNo             = upd[0];
    this.propertyGroupName  = upd[1];
    this.propertyGroup      = upd[2];
    this.monthAccRent       = upd[3];
    
    for (var i=0;i<upd.length;i++) this.itemPack.push(upd[i]);
  }

  show(){
    var text = `itemRptStatus: \n(itemNo=${this.itemNo}, propertyGroupName=${this.propertyGroupName}, propertyGroup=${this.propertyGroup}, monthAccRent=${this.monthAccRent})`;
    return text;
  };

}

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
  report_analysis();
  report_status();
  report_event();
}

function report_analysis() {
  /////////////////////////////////////////
  // README
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const topRowOfs = 1; // the offset from the top row, A2 is 1
  
  /////////////////////////////////////////
  // Clear sheet
  /////////////////////////////////////////
  var pos = new itemRptAnalysis([]);
  SheetRptAnalysisName.getRange(1+topRowOfs,pos.ColPos_ItemNo,SheetRptAnalysisName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  SheetRptAnalysisName.getRange(1          ,pos.ColPos_MonthAccRent,SheetRptAnalysisName.getLastRow(),CFG_Val_obj["CFG_MonthAccRent_NUM"]).clear(); // clear MonthAccRent column including header
  
  /////////////////////////////////////////
  // Cal month acc rent
  /////////////////////////////////////////
  var data = SheetRptAnalysisName.getRange(1+topRowOfs,1,SheetRptAnalysisName.getLastRow()-topRowOfs,SheetRptAnalysisName.getLastColumn()).getValues();
  for(i=0;i<data.length;i++){
    var itemNo = i;

    var propertyGroup_regex = data[i][pos.ColPos_PropertyGroup-1].toString().replace(/[\s|\n|\r|\t]/g,"");
    var srhGroup_arr = propertyGroup_regex.split(";");
    
    var stDate = new Date(CONST_TODAY_DATE.getTime());
    stDate.setDate(1);
    stDate.setMonth(CONST_TODAY_DATE.getMonth()); 
    var edDate = new Date(CONST_TODAY_DATE.getTime());
    edDate.setDate(1);
    edDate.setMonth(CONST_TODAY_DATE.getMonth()+1); 
    for (j=0;j<CFG_Val_obj["CFG_MonthAccRent_NUM"];j++){
      var accRent = 0;
      var accRentDetails_arr = new Array();
      for (k=0;k<srhGroup_arr.length;k++){
        var srhPtn = "^" + srhGroup_arr[k].toString().replace(/[*]/g,"[\u4E00-\uFF5A0-9A-Za-z\u0020-\u007E]?") + "$";
        var regExp = new RegExp(srhPtn,"gi");
        // from BankRecord
        for (var kk=0;kk<GLB_BankRecord_arr.length;kk++) {
          var record =new itemBankRecord(GLB_BankRecord_arr[kk]);
          if (record.rentProperty != null) {
            if ((stDate <= record.date) && (record.date < edDate)) {
              if (record.rentProperty.toString().match(regExp) != null) {
                accRentDetails_arr.push(`Record: ${record.itemNo}\t${record.rentProperty}\t${record.amount}\n`);
                accRent += record.amount;
              }
            }
          }
        }

        // from MiscCost Cash_Rent
        for (var kk=0;kk<GLB_MiscCost_arr.length;kk++) {
          var misc =new itemMiscCost(GLB_MiscCost_arr[kk]);
          if ((misc.rentProperty != null) && (misc.type == misc.MiscType_CashRent)) {
            if ((stDate <= misc.date) && (misc.date < edDate)) {
              if (misc.rentProperty.toString().match(regExp) != null) {
                accRentDetails_arr.push(`MISC: ${misc.itemNo}\t${misc.rentProperty}\t${misc.amount}\n`);
                accRent += misc.amount;
              }
            }
          }
        }
      }
      SheetRptAnalysisName.getRange(0+topRowOfs,pos.ColPos_MonthAccRent+j).setValue(`${stDate.getMonth()+1}/${stDate.getFullYear()}`);
      SheetRptAnalysisName.getRange(1+topRowOfs+i,pos.ColPos_MonthAccRent+j).setValue(accRent).setNote(`${accRentDetails_arr}`);
      stDate.setMonth(stDate.getMonth()-1); 
      edDate.setMonth(edDate.getMonth()-1);

      
    }
    SheetRptAnalysisName.getRange(1+topRowOfs+i,pos.ColPos_ItemNo).setValue(i);
    
  }
  
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
        
        var rentDay = new Date(contract.fromDate);
        if (CONST_TODAY_DATE.getDate() >= contract.fromDate.getDate()) rentDay.setMonth(CONST_TODAY_DATE.getMonth()+1);
        else rentDay.setMonth(CONST_TODAY_DATE.getMonth());
        rentDay.setFullYear(CONST_TODAY_DATE.getFullYear());
        var rentDayDist = Math.floor((rentDay - CONST_TODAY_DATE) / CONST_MILLIS_PER_DAY);
        
        var status = `6.Rent arear is -${rptNum}, ${Math.floor(rentArrear*10/property.curRent)/10} month, ${rentDayDist} due days.`;
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
    if ((contract.validContract || CFG_Val_obj["CFG_ReportEvent_ShowEndContract"]) && (contract.fromDate <= CONST_TODAY_DATE)) {
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
        var match = false;
        if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == contract.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"")) { 
          // for manually assign contract No record
          var fromDate = new Date(contract.fromDate.getTime()-CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_FromDateMargin"]);
          var toDate   = new Date(finishDate.getTime()       +CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_ToDateMargin"]);
          if ((fromDate <= record.date) && (record.date < toDate)){
            match = true;
          }
        }
        else {
          // for normal case
          if ((contract.fromDate.getTime() <= record.date) && (record.date < finishDate)){
            // for account search
            if (contract.tenantAccount_arr.indexOf(record.fromAccount.toString())!=-1) {
              match = true;
            }
            // for account name search
            else if (contract.tenantAccountName_regex.replace(/[\s|\n|\r|\t]/g,"")!='') {
              var accountName_arr = contract.tenantAccountName_regex.replace(/[\s|\n|\r|\t]/g,"").split(";");
              for (k=0;k<accountName_arr.length;k++){
                var srhPtn = "^" + accountName_arr[k].toString().replace(/[*]/g,"[\u4E00-\uFF5A0-9A-Za-z\u0020-\u007E]?") + "$";
                var regExp = new RegExp(srhPtn,"gi");
                var fromAccountName_arr = record.fromAccountName.toString().replace(/[\s|\n|\r|\t]/g,"").split(";");
                for (kk=0;kk<fromAccountName_arr.length;kk++){
                  if (fromAccountName_arr[kk].toString().match(regExp) != null){
                    match = true;
                  }
                }
              }
            }
          }
        }
        
        if (match) {
          var item    = new itemRptEvent([]);
          var date    = Utilities.formatDate(record.date, "GMT+8", "yyyy/MM/dd");
          var event   = `5. Bank record.`;
          var amount  = record.amount;
          var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
          item.update(upd);
          GLB_RptEvent_arr.push(item.itemPack);
          itemNo += 1;
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
