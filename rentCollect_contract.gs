const CONST_TODAY_DATE = new Date();
const CONST_MILLIS_PER_DAY = 1000 * 60 * 60 * 24;


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
    
    var expect_rent = ((target_date.getFullYear()-item.fromDate.getFullYear())*12 + (target_date.getMonth()-item.fromDate.getMonth()) + (target_date.getDate() > item.fromDate.getDate())) * item.amount; // payment increased when the (date of TODAY_DATE) >= (date of item.fromDate)
    // Logger.log(`expect_rent=${expect_rent}, item.fromDate.getDate()=${item.fromDate.getDate()}. ${item.show()}`);
    
    var expect_util = 0;
    for (j=0;j<GLB_UtilBill_arr.length;j++) {
      var util = new itemUtilBill(GLB_UtilBill_arr[j]);
      if ((item.fromDate <= util.date) && (util.date < item.toDate) && (util.rentProperty == item.rentProperty)){
        expect_util += util.amount;
      }
    }

    var expect_misc = 0;
    for (j=0;j<GLB_MiscCost_arr.length;j++) {
      var misc = new itemUtilBill(GLB_MiscCost_arr[j]);
      if ((item.fromDate <= misc.date) && (misc.date < item.toDate) && (misc.rentProperty == item.rentProperty)){
        if (misc.type == misc.MiscType_Charge_Fee)      expect_misc += misc.amount;
        else if (misc.type == misc.MiscType_CashRent)   expect_misc += misc.amount;
        else if (misc.type == misc.MiscType_Repare_Fee) expect_misc += 0; // expect to be absorbed by preperty owner
        else if (misc.type == misc.MiscType_Refund)     expect_misc -= misc.amount; // return back to tenant
        else {var errMsg = `[rentCollect_contract] Type of MiscNo: ${item.itemNo} is invalid!`; reportErrMsg(errMsg);}
      }
    }
    
    var actual_payemnt = 0;
    for (var j=0;j<GLB_BankRecord_arr.length;j++) {
      var record =new itemBankRecord(GLB_BankRecord_arr[j]);
      // record.show();
      // Logger.log(`AAA ${item.tenantAccount_arr.indexOf(record.fromAccount)}, BBB ${record.fromAccount}`)
      var finishDate;
      if (item.endContract) finishDate = item.endDate;
      else finishDate = item.toDate;
      
      if ((item.fromDate <= record.date) && (record.date < finishDate) && (item.tenantAccount_arr.indexOf(record.fromAccount)!=-1)) {
        // Logger.log(`hit record: ${record.show()}`);
        if (record.contractNo == null) { // first record update
          actual_payemnt += record.amount;
          SheetBankRecordName.getRange(1+topRowOfs+j,record.ColPos_ContractNo).setValue(item.itemNo);
          SheetBankRecordName.getRange(1+topRowOfs+j,record.ColPos_rentProperty).setValue(item.rentProperty);
          var upd = [item.itemNo,item.rentProperty];
          record.update(upd);
        }
        else if (record.contractNo == item.itemNo) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} How come to overlap contractNo??`; reportErrMsg(errMsg);}
        else if (record.rentProperty == item.rentProperty) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} How come to overlap rentProperty??`; reportErrMsg(errMsg);}
      }
    }
    
    var rentArrear = actual_payemnt - expect_rent - expect_util - expect_misc;
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
                            ((item.fromDate <= CONST_TODAY_DATE) && (CONST_TODAY_DATE <  item.toDate)                                   ) || 
                            (                                       (CONST_TODAY_DATE >= item.toDate) && ((rentArrear+item.deposit) != 0));
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
    // check end contract date
    /////////////////////////////////////////
    // var endDate = SheetContractName.getRange(1+topRowOfs+j,item.ColPos_EndDate).getValues();
    endDate = item.endDate;
    // Logger.log(`FFF: ${item.endDate}`);
    if (item.endContract) { // to update the endDate
      if (endDate.toString().replace(" ","") == "") {
        if (CONST_TODAY_DATE < item.fromDate) {
          if (1) {var errMsg = `[rentCollect_contract] ContractNo: ${item.itemNo} EndContract incorrect`; reportErrMsg(errMsg);}
        }
        else if (CONST_TODAY_DATE < item.toDate) {
          var endDate = CONST_TODAY_DATE;
          SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(endDate); // abnormal end date, terminate contract ahead.
        }
        else if (item.toDate <= CONST_TODAY_DATE) {
          var endDate = item.toDate;
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
      SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(""); 
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
  for (i=0;i<GLB_Property_arr.length;i++){
    var item = new itemProperty(GLB_Property_arr[i]);
    var isNotLinkContract = true;
    // Logger.log(`111: ${item.show()}`);
    for(j=0;j<GLB_Contract_arr.length;j++) {
      var contract = new itemContract(GLB_Contract_arr[j]);
      var occupied = false;
      // var isValidContract = SheetContractName.getRange(1+topRowOfs+j,contract.ColPos_ValidContract).getValue();
      // Logger.log(`222: ${contract.show()}, isValidContract: ${isValidContract}`);
      if (contract.validContract) {
        if ((item.rentProperty == contract.rentProperty)){
          // since the contract integrity guarantee only one rentProperty is occupied a time, wii hit contract only once.
          if (isNotLinkContract == false) {var errMsg = `[rentCollect_contract] contractNo: ${contract.itemNo} re-link to to rentProperty: ${item.rentProperty}`; reportErrMsg(errMsg);} 
          
          if ((contract.fromDate <= CONST_TODAY_DATE) && (CONST_TODAY_DATE < contract.toDate)) {
            occupied = true;
          }
          SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_DayRest).setValue(contract.dayRest);
          SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_CurRent).setValue(contract.amount);
          SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_ContractNo).setValue(contract.itemNo);
          SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_TenantName).setValue(contract.tenantName);
          SheetPropertyName.getRange(1+topRowOfs+i,item.ColPos_ValidContract).setValue(contract.validContract);

          var upd = [occupied,contract.dayRest,contract.amount,contract.itemNo,contract.tenantName,contract.validContract];
          item.update(upd);
          
          isNotLinkContract = false;
        }
      }
    }

    if (isNotLinkContract) { // for all contract miss case
      var upd = [false,null,null,null,null,false];
      item.update(upd);
    }

  }
}


