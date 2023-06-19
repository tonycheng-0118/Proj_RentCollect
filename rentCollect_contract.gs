function rentCollect_contract() {
  /////////////////////////////////////////
  // README
  // GLB_Tenant_obj is a hash of telant->account
  // itemContract_arr is a collection of all contract link to GLB_Tenant_obj
  // itemContract_arr.apply(chkContractSts()) to show the rent collection readiness.
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  const topRowOfs = 1;

  /////////////////////////////////////////
  // Clear sheet
  /////////////////////////////////////////
  var record =new itemBankRecord([]);
  if (SheetBankRecordName.getLastRow()>1) SheetBankRecordName.getRange(1+topRowOfs,record.ColPos_ContractNo,SheetBankRecordName.getLastRow()-topRowOfs,SheetBankRecordName.getLastColumn()).clear();
  
  /////////////////////////////////////////
  // Update Status of each contract
  /////////////////////////////////////////
  for(i=0;i<GLB_Contract_arr.length;i++){
    var item = new itemContract(GLB_Contract_arr[i]);
    Logger.log(`contract: ${item.show()}`);

    /////////////////////////////////////////
    // check rent arrear
    /////////////////////////////////////////
    var target_date = Date();
    if (CONST_TODAY_DATE >= item.toDate) target_date = item.toDate;
    else target_date = CONST_TODAY_DATE;

    var total_rent_count = ((item.toDate.getFullYear()-item.fromDate.getFullYear())*12 + (item.toDate.getMonth()-item.fromDate.getMonth()) + (item.toDate.getDate() > item.fromDate.getDate()));
    if (total_rent_count % item.period) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} total_rent_count is invalid, should be multiple of period!`; reportErrMsg(errMsg);}
    
    var expect_rent_count = ((target_date.getFullYear()-item.fromDate.getFullYear())*12 + (target_date.getMonth()-item.fromDate.getMonth()) + (target_date.getDate() > item.fromDate.getDate())) / item.period; // payment increased when the (date of TODAY_DATE) >= (date of item.fromDate)
    
    var expect_rent = Math.ceil(expect_rent_count) * item.amount + item.deposit; // will include deposit since it is part of the contract expect payment
    // Logger.log(`expect_rent=${expect_rent}, item.fromDate.getDate()=${item.fromDate.getDate()}. ${item.show()}`);
    
    var expect_util = 0;
    for (j=0;j<GLB_UtilBill_arr.length;j++) {
      var util = new itemUtilBill(GLB_UtilBill_arr[j]);
      if ((item.fromDate <= util.date) && (util.date < item.toDate) && (util.rentProperty == item.rentProperty)){
        expect_util += util.amount;
      }
    }

    var expect_misc = 0;
    var actual_cash_payment = 0;
    for (j=0;j<GLB_MiscCost_arr.length;j++) {
      var misc = new itemMiscCost(GLB_MiscCost_arr[j]);
      if ((item.fromDate <= misc.date) && (misc.date < item.toDate) && (misc.rentProperty == item.rentProperty)){
        if (misc.type == misc.MiscType_CashRent) actual_cash_payment += misc.amount; // cash paid by the tenant, TODO, find a better way to collect the cash payment.
        else expect_misc += misc.expect_misc();
      }
    }
    
    var actual_transfer_payment = 0;
    for (var j=0;j<GLB_BankRecord_arr.length;j++) {
      var record =new itemBankRecord(GLB_BankRecord_arr[j]);
      // record.show();
      // Logger.log(`AAA ${item.tenantAccount_arr.indexOf(record.fromAccount)}, BBB ${record.fromAccount}`)
      var finishDate;

      if (item.endContract) finishDate = new Date(item.endDate.getTime()+CONST_MILLIS_PER_DAY);
      else if (CONST_TODAY_DATE <= item.toDate) finishDate = new Date(CONST_TODAY_DATE.getTime()+CONST_MILLIS_PER_DAY);
      else finishDate = new Date(item.toDate);
      

      // for contractOverrid case
      var fromDate = new Date(item.fromDate.getTime()-CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_FromDateMargin"]);
      var toDate   = new Date(finishDate.getTime()   +CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_ToDateMargin"]);
      if ((fromDate <= record.date) && (record.date < toDate)){
        if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == item.itemNo.toString().replace(/[\s|\n|\r|\t]/g,"")) { // for manually assign contract No record
          actual_transfer_payment += record.amount;
          SheetBankRecordName.getRange(1+topRowOfs+j,record.ColPos_ContractNo).setValue(item.itemNo);
          SheetBankRecordName.getRange(1+topRowOfs+j,record.ColPos_rentProperty).setValue(item.rentProperty);
          var upd = [item.itemNo,item.rentProperty];
          record.update(upd);
        }
      }
      
      // for normal case
      if ((item.fromDate.getTime() <= record.date) && (record.date < finishDate)){
        if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") == "") {
          if (item.tenantAccount_arr.indexOf(record.fromAccount.toString())!=-1) {
            // Logger.log(`hit record: ${record.show()}`);
            if (record.contractNo == null) { // first record update
              actual_transfer_payment += record.amount;
              SheetBankRecordName.getRange(1+topRowOfs+j,record.ColPos_ContractNo).setValue(item.itemNo);
              SheetBankRecordName.getRange(1+topRowOfs+j,record.ColPos_rentProperty).setValue(item.rentProperty);
              var upd = [item.itemNo,item.rentProperty];
              record.update(upd);
            }
            else if (record.contractNo == item.itemNo) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} How come to overlap contractNo??`; reportErrMsg(errMsg);}
            else if (record.rentProperty == item.rentProperty) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} How come to overlap rentProperty??`; reportErrMsg(errMsg);}
          }
        }
      }
    }
    
    var rentArrear = actual_transfer_payment + actual_cash_payment - expect_rent - expect_util - expect_misc;
    SheetContractName.getRange(1+topRowOfs+i,item.ColPos_RentArrear).setValue(rentArrear);

    /////////////////////////////////////////
    // check ST_ValidContract
    // if any unsettled payment, then the contract still stands, unless manually terminate the contract
    /////////////////////////////////////////
    if (item.endContract){
      var isValidContract = false; // only way to deactivate the contract
    }
    else {
      var isValidContract = (item.endContract == false) || 
                            ((item.fromDate <= CONST_TODAY_DATE) && (CONST_TODAY_DATE <  item.toDate)                      ) || 
                            (                                       (CONST_TODAY_DATE >= item.toDate) && ((rentArrear) != 0));
    }
    // Logger.log(`GGGG ${isValidContract}, JJJJ: ${item.endContract}`);
    SheetContractName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(isValidContract);


    /////////////////////////////////////////
    // check days rest
    /////////////////////////////////////////
    if (isValidContract){
      var rest = Math.floor((item.toDate - CONST_TODAY_DATE) / CONST_MILLIS_PER_DAY);
      SheetContractName.getRange(1+topRowOfs+i,item.ColPos_DayRest).setValue(rest);
    }
    else {
      SheetContractName.getRange(1+topRowOfs+i,item.ColPos_DayRest).setValue(`N/A`);
    }

    /////////////////////////////////////////
    // update item
    /////////////////////////////////////////
    var upd = [isValidContract,rentArrear,rest];
    item.update(upd);
    // Logger.log(`update contract: ${item.show()}`);

  }

  /////////////////////////////////////////
  // link to BankRecord 
  /////////////////////////////////////////


  /////////////////////////////////////////
  // link to UtilBill 
  /////////////////////////////////////////
  var isNotLinkContract = true;
  var pos = new itemUtilBill(GLB_UtilBill_arr[0]);
  SheetUtilBillName.getRange(1+topRowOfs,pos.ColPos_ContractNo,GLB_UtilBill_arr.length,1).setValue('No Matched');
  SheetUtilBillName.getRange(1+topRowOfs,pos.ColPos_TenantName,GLB_UtilBill_arr.length,1).setValue('No Matched');
  SheetUtilBillName.getRange(1+topRowOfs,pos.ColPos_ValidContract,GLB_UtilBill_arr.length,1).setValue('No Matched');
  for (i=0;i<GLB_UtilBill_arr.length;i++){
    var item = new itemUtilBill(GLB_UtilBill_arr[i]);
    // item.show();
    for(j=0;j<GLB_Contract_arr.length;j++){
      var contract = new itemContract(GLB_Contract_arr[j]);
      // contract.show();
      if ((contract.fromDate <= item.date) && (item.date < contract.toDate) && (item.rentProperty == contract.rentProperty)) {
        // var isValidContract = SheetContractName.getRange(1+topRowOfs+j,contract.ColPos_ValidContract).getValues();
        SheetUtilBillName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contract.itemNo);
        SheetUtilBillName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(contract.tenantName);
        SheetUtilBillName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(contract.validContract);
        
        var upd = [contract.itemNo,contract.tenantName,contract.validContract];
        item.update(upd);
        
        isNotLinkContract = false;
        break;
      }
    }
    if (isNotLinkContract) {var errMsg = `[rentCollect_contract] UtilBillNo: ${item.itemNo} cannot link to any ContractNo`; reportErrMsg(errMsg);}
  }

  /////////////////////////////////////////
  // link to MiscCost 
  /////////////////////////////////////////
  var isNotLinkContract = true;
  var pos = new itemMiscCost(GLB_MiscCost_arr[0])
  SheetMiscCostName.getRange(1+topRowOfs,pos.ColPos_ContractNo,GLB_MiscCost_arr.length,1).setValue('No Matched');
  SheetMiscCostName.getRange(1+topRowOfs,pos.ColPos_TenantName,GLB_MiscCost_arr.length,1).setValue('No Matched');
  SheetMiscCostName.getRange(1+topRowOfs,pos.ColPos_ValidContract,GLB_MiscCost_arr.length,1).setValue('No Matched')
  for (i=0;i<GLB_MiscCost_arr.length;i++){
    var item = new itemMiscCost(GLB_MiscCost_arr[i]);
    // item.show();
    for(j=0;j<GLB_Contract_arr.length;j++){
      var contract = new itemContract(GLB_Contract_arr[j]);
      // contract.show();
      if ((contract.fromDate <= item.date) && (item.date < contract.toDate) && (item.rentProperty == contract.rentProperty)) {
        // var isValidContract = SheetContractName.getRange(1+topRowOfs+i,contract.ColPos_ValidContract).getValue();
        SheetMiscCostName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contract.itemNo);
        SheetMiscCostName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(contract.tenantName);
        SheetMiscCostName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(contract.validContract);
        
        var upd = [contract.itemNo,contract.tenantName,contract.validContract];
        item.update(upd);
        
        isNotLinkContract = false;
        break;
      }
    }
    if (isNotLinkContract) {var errMsg = `[rentCollect_contract] MiscCostNo: ${item.itemNo} cannot link to any ContractNo`; reportErrMsg(errMsg);}
  }

  /////////////////////////////////////////
  // link to Property 
  /////////////////////////////////////////
  var pos = new itemProperty(GLB_Property_arr[0]);
  SheetPropertyName.getRange(1+topRowOfs,pos.ColPos_Occupied,GLB_Property_arr.length,1).setValue(false);
  SheetPropertyName.getRange(1+topRowOfs,pos.ColPos_DayRest,GLB_Property_arr.length,1).setValue(`N/A`);
  SheetPropertyName.getRange(1+topRowOfs,pos.ColPos_CurRent,GLB_Property_arr.length,1).setValue(`N/A`);
  SheetPropertyName.getRange(1+topRowOfs,pos.ColPos_ContractNo,GLB_Property_arr.length,1).setValue(`N/A`);
  SheetPropertyName.getRange(1+topRowOfs,pos.ColPos_TenantName,GLB_Property_arr.length,1).setValue(`N/A`);
  SheetPropertyName.getRange(1+topRowOfs,pos.ColPos_ValidContract,GLB_Property_arr.length,1).setValue(false);
  for (var i=0;i<GLB_Property_arr.length;i++){
    var item = new itemProperty(GLB_Property_arr[i]);
    var isNotLinkContract = true;
    var occupied = false;
    var linkContractNo = -1;
    var oldestValidContractDate = CONST_SuperFeatureDate;
    // Logger.log(`111: i: ${i}, ${item.show()}`);
    for(j=0;j<GLB_Contract_arr.length;j++) {
      var contract = new itemContract(GLB_Contract_arr[j]);
      // var isValidContract = SheetContractName.getRange(1+topRowOfs+j,contract.ColPos_ValidContract).getValue();
      // Logger.log(`222: ${contract.show()}, isValidContract: ${isValidContract}`);
      if (contract.validContract) {
        if ((item.rentProperty == contract.rentProperty)){
          // since the contract integrity guarantee only one rentProperty is occupied a time, wii hit contract only once.
          // But if a futurn increamental contract will make multi contract mapping to same property.
          // if (isNotLinkContract == false) {var errMsg = `[rentCollect_contract] contractNo: ${contract.itemNo} re-link to to rentProperty: ${item.rentProperty}`; reportErrMsg(errMsg);} 
          
          if ((contract.fromDate <= CONST_TODAY_DATE) && (CONST_TODAY_DATE < contract.toDate)) occupied = true;
          
          if (oldestValidContractDate > contract.fromDate) {
            oldestValidContractDate = new Date(contract.fromDate.getTime());
            linkContractNo = contract.itemNo; // the property.contractNo will link the oldest valid contract.
          }

          isNotLinkContract = false;
          
        }
      }
    }

    if (isNotLinkContract) { // for all contract miss case
      var upd = [false,null,null,null,null,false];
      item.update(upd);
    }
    else {
      var contract = new itemContract(GLB_Contract_arr[findContractNoPos(linkContractNo)]);
      SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_DayRest).setValue(contract.dayRest);
      SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_CurRent).setValue(Math.floor(contract.amount/contract.period));
      SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contract.itemNo);
      SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(contract.tenantName);
      SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(contract.validContract);

      var upd = [occupied,contract.dayRest,contract.amount,contract.itemNo,contract.tenantName,contract.validContract];
      item.update(upd);
    }

  }
}


