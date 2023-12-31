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
  SheetRptAnalysisName.getRange(1+topRowOfs,pos.ColPos_ItemNo,SheetRptAnalysisName.getLastRow()-topRowOfs,1).clear({ formatOnly: false, contentsOnly: true }); // clear itemNo column
  SheetRptAnalysisName.getRange(1          ,pos.ColPos_MonthAccRent,SheetRptAnalysisName.getLastRow(),CFG_Val_obj["CFG_MonthAccRent_NUM"]).clear({ formatOnly: false, contentsOnly: true }); // clear MonthAccRent column including header
  
  /////////////////////////////////////////
  // Cal month acc rent
  /////////////////////////////////////////
  var data = SheetRptAnalysisName.getRange(1+topRowOfs,1,SheetRptAnalysisName.getLastRow()-topRowOfs,SheetRptAnalysisName.getLastColumn()).getValues();
  for(i=0;i<data.length;i++){
    var itemNo = i;

    var propertyGroup_regex = data[i][pos.ColPos_PropertyGroup-1].toString().replace(/[\s|\n|\r|\t]/g,"");
    var srhGroup_arr = propertyGroup_regex.split(";");
    var propertyExclude_regex = data[i][pos.ColPos_PropertyExclude-1].toString().replace(/[\s|\n|\r|\t]/g,"");
    var srhExclude_arr = propertyExclude_regex.split(";");
    
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
        var srhMatchPtn = "^" + srhGroup_arr[k].toString().replace(/[*]/g,"[\u4E00-\uFF5A0-9A-Za-z\u0020-\u007E]?") + "$";
        var srhMatchRegexp = new RegExp(srhMatchPtn,"gi");
        // from BankRecord
        for (var kk=0;kk<GLB_BankRecord_arr.length;kk++) {
          var record =new itemBankRecord(GLB_BankRecord_arr[kk]);
          if (record.rentProperty != null) {
            var isExclude = 0;
            // search exclude property
            for (kkk=0;kkk<srhExclude_arr.length;kkk++){
              var srhExcludePtn = "^" + srhExclude_arr[kkk].toString() + "$";
              var srhExcludeRegexp = new RegExp(srhExcludePtn,"gi");
              if (record.rentProperty.toString().match(srhExcludeRegexp) != null) {
                isExclude = 1;
              }
            }
            
            // search match property
            if ((isExclude==0) && (stDate <= record.date) && (record.date < edDate)) {
              if (record.rentProperty.toString().match(srhMatchRegexp) != null) {
                accRentDetails_arr.push(`Record: ${record.itemNo}\t${Utilities.formatDate(record.date, 'GMT+8', 'yyyy/MM/dd')}\t${record.rentProperty}\t${record.amount}\n`);
                accRent += record.amount;
              }
            }
          }
        }

        // from MiscCost Cash_Rent
        for (var kk=0;kk<GLB_MiscCost_arr.length;kk++) {
          var misc =new itemMiscCost(GLB_MiscCost_arr[kk]);
          if ((misc.rentProperty != null) && (misc.type == misc.MiscType_CashRent)) {
            var isExclude = 0;
            // search exclude property
            for (kkk=0;kkk<srhExclude_arr.length;kkk++){
              var srhExcludePtn = "^" + srhExclude_arr[kkk].toString() + "$";
              var srhExcludeRegexp = new RegExp(srhExcludePtn,"gi");
              if (misc.rentProperty.toString().match(srhExcludeRegexp) != null) {
                isExclude = 1;
              }
            }

            // search match property
            if ((isExclude==0) && (stDate <= misc.date) && (misc.date < edDate)) {
              if (misc.rentProperty.toString().match(srhMatchRegexp) != null) {
                accRentDetails_arr.push(`MISC: ${misc.itemNo}\t${Utilities.formatDate(misc.date, 'GMT+8', 'yyyy/MM/dd')}\t${misc.rentProperty}\t${misc.amount}\n`);
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
  if (SheetRptStatusName.getLastRow()>1) SheetRptStatusName.getRange(1+topRowOfs,1,SheetRptStatusName.getLastRow()-topRowOfs,SheetRptStatusName.getLastColumn()).clear({ formatOnly: false, contentsOnly: true }); // in case of empty
  for (var i=0;i<GLB_Property_arr.length;i++){
    var item = new itemRptStatus([]);
    var property = new itemProperty(GLB_Property_arr[i]);
    // Logger.log(`property for status: ${property.show()}`);

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
    
    
    
    // SheetRptStatusName.getRange(1+topRowOfs+i,1).setValue(itemNo);
    // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_RentProperty).setValue(rentProperty);
    // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Occupied).setValue(occupied);
    // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(isValidContract);
    // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_AccBalance).setValue(accBalance);
    
    // Link the information from sheets
    // if (occupied) { // if not occupied, left the blank along
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(tenantName);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_RentArrear).setValue(rentArrear);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Deposit).setValue(deposit);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_CurRent).setValue(curRent);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_DayRest).setValue(dayRest);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contractNo);
    // }
    // else if (isValidContract) { // may have end of cnstract but with unpaid rents case.
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(tenantName);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_RentArrear).setValue(rentArrear);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Deposit).setValue(deposit);
    //   SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contractNo);
    // }
    
    // Logger.log(`contract: ${contract.show()}`);
    // report status with color
    if ((property.occupied == false) && (property.validContract == false)) {
      var status = "9.Vacancy";
      // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
      // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Grey);
    } else if ((property.occupied == false) && (property.validContract == true)) {
      if ((rentArrear) > 0) {
        var rptNum = (1000000 + rentArrear).toPrecision(7).toString().substring(1); // for add leading zero
        var status = `5.Need to refund by ${rptNum}.`;
        // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Yellow);
      } else if ((rentArrear) < 0) {
        var rptNum = (1000000 + Math.abs(rentArrear)).toPrecision(7).toString().substring(1); // for add leading zero
        var status = `8.Need to charge by ${rptNum}.`;
        // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Red);
      } else if (CONST_TODAY_DATE < contract.fromDate) {
        var status = `1.Contract is not started yet.`;
      } else {
        var status = `7.Need to final check.`;
        // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Yellow);
      }
    } else if ((property.occupied == true) && (property.validContract == true)) {
      if (rentArrear + property.curRent*CFG_Val_obj["CFG_ReportArrearMargin"] < 0) {
        var rptNum = (1000000 + Math.abs(rentArrear)).toPrecision(7).toString().substring(1); // for add leading zero
        
        // var rentDay = new Date();
        // rentDay.setSeconds(0);
        // rentDay.setMinutes(0);
        // rentDay.setHours(0);
        // rentDay.setDate(contract.fromDate.getDate());
        // rentDay.setMonth(CONST_TODAY_DATE.getMonth());
        // rentDay.setFullYear(CONST_TODAY_DATE.getFullYear());
        // if (rentDay < CONST_TODAY_DATE) rentDay += CONST_MILLIS_PER_DAY*30; // for estimated rentDayDist, varied by diff month
        // var rentDayDist = Math.floor((rentDay - CONST_TODAY_DATE) / CONST_MILLIS_PER_DAY);
        var rentDayDist = 0;
        var rentDay = contract.fromDate.getDate();
        if (CONST_TODAY_DATE.getDate() < rentDay) {
          rentDayDist = CONST_TODAY_DATE.getDate() + 30 - rentDay;
        } else {
          rentDayDist = CONST_TODAY_DATE.getDate()      - rentDay;
        }
        
        var status = `6.Rent arear is -${rptNum}, ${Math.floor(rentArrear*10/property.curRent)/10} month, ${rentDayDist} due days.`;
        // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Red);
      }
      else if (property.occupied && property.dayRest <= 30) {
        var status = `4.Contract near end within ${property.dayRest} days.`;
        // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Yellow);
      }
      else {
        var status = "0.Good!!!";
        // SheetRptStatusName.getRange(1+topRowOfs+i,item.ColPos_Status).setValue(status);
        // SheetRptStatusName.getRange(1+topRowOfs+i,1,1,SheetRptStatusName.getLastColumn()).setBackground(Color_Green);
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
  
  // show on the sheet
  var item    = new itemRptStatus([]);
  var range = SheetRptStatusName.getRange(1+topRowOfs,1,GLB_RptStatus_arr.length,item.itemPackMaxLen).setValues(GLB_RptStatus_arr);
  
  // sort with priority
  var range = SheetRptStatusName.getRange(1+topRowOfs,1,SheetRptStatusName.getLastRow()-topRowOfs,SheetRptStatusName.getLastColumn());
  range.sort({column:item.ColPos_RentProperty,ascending: true});
  range.sort({column:item.ColPos_RentArrear,ascending: true});
  range.sort({column:item.ColPos_Status,ascending: false});
  
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
  if (SheetRptEventName.getLastRow()>1) SheetRptEventName.getRange(1+topRowOfs,1,SheetRptEventName.getLastRow()-topRowOfs,SheetRptEventName.getLastColumn()).clear({ formatOnly: false, contentsOnly: true }); // in case of empty
  
  var itemNo = 0;

  for (var i=0;i<GLB_Contract_arr.length;i++){
    var contract = new itemContract(GLB_Contract_arr[i]);
    Logger.log(`contract for event: ${contract.show()}`)
    if ((contract.validContract || CFG_Val_obj["CFG_ReportEvent_ShowEndContract"]) && (contract.fromDate <= CONST_TODAY_DATE)) {
      // Event: start contract with deposit
      var fromDate= Utilities.formatDate(contract.fromDate, "GMT+8", "yyyy/MM/dd");
      var toDate  = Utilities.formatDate(contract.toDate, "GMT+8", "yyyy/MM/dd");
      var event   = `1. Start of the contract from ${fromDate} to ${toDate} with deposit = ${contract.deposit}`;
      var amount  = -1 * contract.deposit;
      setRptEventItem(fromDate,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount);

      // Event: rent bill
      var finishDate;
      if (contract.toDate <= CONST_TODAY_DATE) finishDate = contract.toDate;
      else finishDate = new Date(CONST_TODAY_DATE.getTime()+CONST_MILLIS_PER_DAY);
      
      for (var j = new Date(contract.fromDate);j < finishDate;) {
        var item    = new itemRptEvent([]);
        var date    = Utilities.formatDate(j, "GMT+8", "yyyy/MM/dd");
        var event   = `2. Rent bill is ${contract.amount}`;
        var amount  = -1 * contract.amount;
        setRptEventItem(date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount);

        // next month
        j.setMonth(j.getMonth() + contract.period);
      }

      // Event: util bill
      // for (var j=0;j<GLB_UtilBill_arr.length;j++) {
      //   var util = new itemUtilBill(GLB_UtilBill_arr[j]);
      //   var match = false;

      //   if (util.rentProperty == contract.rentProperty) {
      //     if (util.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == contract.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"")) { 
      //       // for manually assign contract No record
      //       if (isContractLegalDateRange(new Date(util.date),contract,true)) {
      //         match = true;
      //       }
      //     } else if (util.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
      //       if (isContractLegalDateRange(new Date(util.date),contract,false)) {
      //         match = true;
      //       }
      //     }

      //     if (match) {
      //       var item    = new itemRptEvent([]);
      //       var date    = Utilities.formatDate(util.date, "GMT+8", "yyyy/MM/dd");
      //       var event   = `3. Util bill is ${util.amount}`;
      //       var amount  = -1 * util.amount;
      //       var upd     = [itemNo,util.date,util.rentProperty,contract.tenantName,contract.itemNo,event,amount];
      //       item.update(upd);
      //       // SheetRptEventName.appendRow(item.itemPack);
      //       GLB_RptEvent_arr.push(item.itemPack);
      //       itemNo += 1;
      //     }
      //   }
      // }
      
      // Event: misc bill
      // for (var j=0;j<GLB_MiscCost_arr.length;j++) {
      //   var misc = new itemMiscCost(GLB_MiscCost_arr[j]);
      //   var match = false;
      //   if (misc.rentProperty == contract.rentProperty) {
      //     if (misc.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == contract.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"")) { 
      //       // for manually assign contract No record
      //       if (isContractLegalDateRange(new Date(misc.date),contract,true)) {
      //         match = true;
      //       }
      //     } else if (misc.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
      //       if (isContractLegalDateRange(new Date(misc.date),contract,false)) {
      //         match = true;
      //       }
      //     }

      //     if (match) {
      //       var item    = new itemRptEvent([]);
      //       var date    = Utilities.formatDate(misc.date, "GMT+8", "yyyy/MM/dd");
      //       var event   = `4. Misc cost, type = ${misc.type}.`;
      //       var amount  = -1 * misc.expect_misc();

      //       var upd     = [itemNo,date,misc.rentProperty,contract.tenantName,contract.itemNo,event,amount];
      //       item.update(upd);
      //       // SheetRptEventName.appendRow(item.itemPack);
      //       GLB_RptEvent_arr.push(item.itemPack);
      //       itemNo += 1;
      //     }
      //   }
      // }
      
      // Event: bank record
      // for (var j=0;j<GLB_BankRecord_arr.length;j++) {
      //   var record = new itemBankRecord(GLB_BankRecord_arr[j]);
      //   var match = false;
      //   if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == contract.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"")) { 
      //     // for manually assign contract No record
      //     if (isContractLegalDateRange(new Date(record.date),contract,true)) {
      //       match = true;
      //     }
      //   } else if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
      //     // for normal case
      //     if (isContractLegalDateRange(new Date(record.date),contract,false)){
      //       // for account search
      //       if (contract.tenantAccount_arr.indexOf(record.fromAccount.toString())!=-1) {
      //         match = true;
      //       }
      //       // for account name search
      //       else if (contract.tenantAccountName_regex.replace(/[\s|\n|\r|\t]/g,"")!='') {
      //         var accountName_arr = contract.tenantAccountName_regex.replace(/[\s|\n|\r|\t]/g,"").split(";");
      //         for (k=0;k<accountName_arr.length;k++){
      //           var srhPtn = "^" + accountName_arr[k].toString().replace(/[*]/g,"[\u4E00-\uFF5A0-9A-Za-z\u0020-\u007E]?") + "$";
      //           var regExp = new RegExp(srhPtn,"gi");
      //           var fromAccountName_arr = record.fromAccountName.toString().replace(/[\s|\n|\r|\t]/g,"").split(";");
      //           for (kk=0;kk<fromAccountName_arr.length;kk++){
      //             if (fromAccountName_arr[kk].toString().match(regExp) != null){
      //               match = true;
      //             }
      //           }
      //         }
      //       }
      //     }
      //   }
        
      //   if (match) {
      //     var item    = new itemRptEvent([]);
      //     var date    = Utilities.formatDate(record.date, "GMT+8", "yyyy/MM/dd");
      //     var event   = `5. Bank record.`;
      //     var amount  = record.amount;
      //     var upd     = [itemNo,date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount];
      //     item.update(upd);
      //     GLB_RptEvent_arr.push(item.itemPack);
      //     itemNo += 1;
      //   }
      // }

      // Event: The contract is over but not endContract
      if (contract.toDate <= CONST_TODAY_DATE) {
        var date    = Utilities.formatDate(contract.toDate, "GMT+8", "yyyy/MM/dd");
        var event   = `6. Contract exceed last data ${date}, but not endContract!`;
        var amount  = null;
        setRptEventItem(date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount);
      }

      // Event: The latest status
      for (var j=0;j<GLB_RptStatus_arr.length;j++) {
        var rptStatus = new itemRptStatus(GLB_RptStatus_arr[j]);
        if (rptStatus.rentProperty == contract.rentProperty) {
          var date    = Utilities.formatDate(CONST_TODAY_DATE, "GMT+8", "yyyy/MM/dd");
          var event   = `9. Current stauts: ${rptStatus.status}`;
          var amount  = null;
          setRptEventItem(date,contract.rentProperty,contract.tenantName,contract.itemNo,event,amount);
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

function setRptEventItem(date,rentProperty,tenantName,contractNo,event,amount) {
  var item    = new itemRptEvent([]);
  var upd     = [GLB_rptEventItemNo,date,rentProperty,tenantName,contractNo,event,amount];
  item.update(upd);
  GLB_RptEvent_arr.push(item.itemPack);
  GLB_rptEventItemNo++;
}
