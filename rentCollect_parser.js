/////////////////////////////////////////
// Function
/////////////////////////////////////////
function rentCollect_parser() {
  // rentCollect_parser_LineMsg
  var time_start_parser_LineMsg = new Date();
  rentCollect_parser_LineMsg();
  var time_finish_parser_parser_LineMsg = new Date();
  if (1) {var info = `[rentCollect_main] parser_LineMsg time_exec(s): ${(time_finish_parser_parser_LineMsg.getTime() - time_start_parser_LineMsg.getTime()) / 1000}`; reportInfoMsg(info);}
  
  // rentCollect_parser_Record
  var time_start_parser_Record = new Date();
  rentCollect_parser_Record();
  var time_finish_parser_Record = new Date();
  if (1) {var info = `[rentCollect_main] parser_Record time_exec(s): ${(time_finish_parser_Record.getTime() - time_start_parser_Record.getTime()) / 1000}`; reportInfoMsg(info);}
  
  // rentCollect_parser_Tenant
  var time_start_parser_Tenant = new Date();
  rentCollect_parser_Tenant();
  var time_finish_parser_Tenant = new Date();
  if (1) {var info = `[rentCollect_main] parser_Tenant time_exec(s): ${(time_finish_parser_Tenant.getTime() - time_start_parser_Tenant.getTime()) / 1000}`; reportInfoMsg(info);}
  
  // rentCollect_parser_UtilBill
  var time_start_parser_UtilBill = new Date();
  rentCollect_parser_UtilBill();
  var time_finish_parser_UtilBill = new Date();
  if (1) {var info = `[rentCollect_main] parser_UtilBill time_exec(s): ${(time_finish_parser_UtilBill.getTime() - time_start_parser_UtilBill.getTime()) / 1000}`; reportInfoMsg(info);}

  // rentCollect_parser_MiscCost
  var time_start_parser_MiscCost = new Date();
  rentCollect_parser_MiscCost();
  var time_finish_parser_MiscCost = new Date();
  if (1) {var info = `[rentCollect_main] parser_MiscCost time_exec(s): ${(time_finish_parser_MiscCost.getTime() - time_start_parser_MiscCost.getTime()) / 1000}`; reportInfoMsg(info);}
  
  // rentCollect_parser_Property
  var time_start_parser_Property = new Date();
  rentCollect_parser_Property();
  var time_finish_parser_Property = new Date();
  if (1) {var info = `[rentCollect_main] parser_Property time_exec(s): ${(time_finish_parser_Property.getTime() - time_start_parser_Property.getTime()) / 1000}`; reportInfoMsg(info);}
  
  // rentCollect_parser_Contract
  var time_start_parser_Contract = new Date();
  rentCollect_parser_Contract();
  var time_finish_parser_Contract = new Date();
  if (1) {var info = `[rentCollect_main] parser_Contract time_exec(s): ${(time_finish_parser_Contract.getTime() - time_start_parser_Contract.getTime()) / 1000}`; reportInfoMsg(info);}
  
  // rentCollect_parser_proxyRecord_MinSheng35
  var time_start_parser_proxyRecord_MinSheng35 = new Date();
  rentCollect_parser_proxyRecord_MinSheng35();
  var time_finish_parser_proxyRecord_MinSheng35 = new Date();
  if (1) {var info = `[rentCollect_main] parser_proxyRecord_MinSheng35 time_exec(s): ${(time_finish_parser_proxyRecord_MinSheng35.getTime() - time_start_parser_proxyRecord_MinSheng35.getTime()) / 1000}`; reportInfoMsg(info);}

  // rentCollect_parser_chkDuplicated
  var time_start_parser_chkDuplicated = new Date();
  rentCollect_parser_chkDuplicated();
  var time_finish_parser_chkDuplicated = new Date();
  if (1) {var info = `[rentCollect_main] parser_chkDuplicated time_exec(s): ${(time_finish_parser_chkDuplicated.getTime() - time_start_parser_chkDuplicated.getTime()) / 1000}`; reportInfoMsg(info);}

}

function rentCollect_parser_Record() {
  var isParseValid = true; // will be false if one of isImpoertValid is false
  
  var time_start_parser_Record_load_BankRecord = new Date();
  load_BankRecord();
  var time_finish_parser_Record_load_BankRecord = new Date();
  if (1) {var info = `[rentCollect_main] parser_Record_load_BankRecord time_exec(s): ${(time_finish_parser_Record_load_BankRecord.getTime() - time_start_parser_Record_load_BankRecord.getTime()) / 1000}`; reportInfoMsg(info);}
  
  if (VAR_isNewRecordImport) {
    var time_start_parser_Record_ESUN = new Date();
    isParseValid &= rentCollect_parser_Record_ESUN(); // 玉山銀行
    var time_finish_parser_Record_ESUN = new Date();
    if (1) {var info = `[rentCollect_main] parser_Record_ESUN time_exec(s): ${(time_finish_parser_Record_ESUN.getTime() - time_start_parser_Record_ESUN.getTime()) / 1000}`; reportInfoMsg(info);}

    var time_start_parser_Record_KTB = new Date();
    isParseValid &= rentCollect_parser_Record_KTB(); // 京城銀行
    var time_finish_parser_Record_KTB = new Date();
    if (1) {var info = `[rentCollect_main] parser_Record_KTB time_exec(s): ${(time_finish_parser_Record_KTB.getTime() - time_start_parser_Record_KTB.getTime()) / 1000}`; reportInfoMsg(info);}

    var time_start_parser_Record_CTBC = new Date();
    isParseValid &= rentCollect_parser_Record_CTBC(); // 中國信託
    var time_finish_parser_Record_CTBC = new Date();
    if (1) {var info = `[rentCollect_main] parser_Record_CTBC time_exec(s): ${(time_finish_parser_Record_CTBC.getTime() - time_start_parser_Record_CTBC.getTime()) / 1000}`; reportInfoMsg(info);}
  }
  
  var time_start_post_BankRecord = new Date();
  if (VAR_isNewRecordImport) {
    if (isParseValid) {
      // backup the previous BankRecord
      bankRecordBackUp();
    } else {
      if (1) {var errMsg = `[rentCollect_parser_Record] isParseValid is invalid!`; reportErrMsg(errMsg);}
    }
  }
  post_BankRecord();
  
  var time_finish_post_BankRecord = new Date();
  if (1) {var info = `[rentCollect_main] post_BankRecord time_exec(s): ${(time_finish_post_BankRecord.getTime() - time_start_post_BankRecord.getTime()) / 1000}`; reportInfoMsg(info);}
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
  const importLastRowSub = 3;  // the last row to ignore
  const importColOfs = 0;  // the offset from the left col, A1 is ofs 0
  const importContentLen = 9;
  const records_obj = [["玉山1001968210899","玉山0899","1001968210899"]];
  // for (const [k,v] of Object.entries(records_obj)) {
  //   Logger.log(`key: ${k}, value: ${v[1]}`);
  // }
  
  /////////////////////////////////////////
  // Import from source sheet
  /////////////////////////////////////////
  var isParseValid = true; // will be false if one of isImpoertValid is false
  for (const [k,v] of Object.entries(records_obj)) {
    
    var time_start_parser_Record_ESUN_import = new Date();
    var isImportValid = rentCollect_import(v[0],false,importRowOfs,importLastRowSub);
    isParseValid = isParseValid && isImportValid;
    var time_finish_parser_Record_ESUN_import = new Date();
    if (1) {var info = `[rentCollect_main] parser_Record_ESUN_import time_exec(s): ${(time_finish_parser_Record_ESUN_import.getTime() - time_start_parser_Record_ESUN_import.getTime()) / 1000}`; reportInfoMsg(info);}

    if (isImportValid) {
      var time_start_parser_Record_ESUN_proc = new Date();
      GLB_Import_arr = new Array(); // to empty the same global array
      var data = SheetImportName.getRange(1+importRowOfs, 1+importColOfs, SheetImportName.getLastRow()-importRowOfs-importLastRowSub, importContentLen).getValues();
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
        var amount = ((act == "TransferOut" || act == "Withdraw" || act == "OtherOut") ? data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"") : data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")) * 1;
        // Logger.log("i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);

        // balance
        var balance = (data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"")) * 1;
        
        // Logger.log("CCC: i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);
        // from
        var fromNote = note_mapping("ESUN", i, date, balance, act, [data[i][2].toString(),data[i][6].toString(),data[i][7].toString(),data[i][8].toString()]);
        
        if (act != ""){
          GLB_Import_arr.push([date,act,amount,balance,fromNote[0],fromNote[1]]);
        }
        
        // Logger.log(`date:${date}, act:${act}, amount: ${amount}, balance: ${balance}, accountName: ${fromNote[0]}, account: ${fromNote[1]}`);
      }
      var time_finish_parser_Record_ESUN_proc = new Date();
      if (1) {var info = `[rentCollect_main] parser_Record_ESUN_proc time_exec(s): ${(time_finish_parser_Record_ESUN_proc.getTime() - time_start_parser_Record_ESUN_proc.getTime()) / 1000}`; reportInfoMsg(info);}
      
      var time_start_parser_Record_ESUN_merge = new Date();
      merge_BankRecord(v[1],v[2]);
      var time_finish_parser_Record_ESUN_merge = new Date();
      if (1) {var info = `[rentCollect_main] parser_Record_ESUN_merge time_exec(s): ${(time_finish_parser_Record_ESUN_merge.getTime() - time_start_parser_Record_ESUN_merge.getTime()) / 1000}`; reportInfoMsg(info);}
    }
  }

  return isParseValid;

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
  const importLastRowSub = 0;  // the last row to ignore
  const importColOfs = 0; // the offset from the left col, A1 is ofs 0
  const importContentLen = 8;
  const records_obj = [["台幣交易明細查詢_9859","京城9859","005220099859"],["台幣交易明細查詢_9141","京城9141","005228009141"],["台幣交易明細查詢_8135","京城8135","005228008135"],["台幣交易明細查詢_3835","京城3835","005220113835"],["台幣交易明細查詢_8231","京城8231","005125018231"]];
  // for (const [k,v] of Object.entries(records_obj)) {
  //   Logger.log(`key: ${k}, value: ${v[1]}`);
  // }
  
  /////////////////////////////////////////
  // Import from source sheet
  /////////////////////////////////////////
  var isParseValid = true; // will be false if one of isImpoertValid is false
  for (const [k,v] of Object.entries(records_obj)) {
    
    var isImportValid = rentCollect_import(v[0],newRecordEnable=false,importRowOfs,importLastRowSub);
    isParseValid = isParseValid && isImportValid;

    if (isImportValid) {
      GLB_Import_arr = new Array(); // to empty the same global array
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
        var amount = ((act == "TransferOut" || act == "Withdraw") ? data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"") : data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")) * 1;
        // Logger.log("i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);

        // balance
        var balance = (data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"")) * 1;
        
        // Logger.log("CCC: i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);
        // from
        var fromNote = note_mapping("KTB", i, date, balance, act, [data[i][3].toString(),data[i][7].toString()]);
        
        if (act != ""){
          GLB_Import_arr.push([date,act,amount,balance,fromNote[0],fromNote[1]]);
        }
        
        // Logger.log(`date:${date}, act:${act}, amount: ${amount}, balance: ${balance}, accountName: ${fromNote[0]}, account: ${fromNote[1]}`);
      }

      merge_BankRecord(v[1],v[2]);
    }
  }
  
  return isParseValid;
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
  const importRowOfs = 1; // the offset from the top row, A1 is ofs 0
  const importLastRowSub = 0;  // the last row to ignore
  
  /////////////////////////////////////////
  // Import from source sheet
  /////////////////////////////////////////
  var isImportValid = rentCollect_import("DEPOSIT_APPLY_RECORD",false,importRowOfs,importLastRowSub);
  var isParseValid = isImportValid;

  /////////////////////////////////////////
  // Parse to itemRecord
  /////////////////////////////////////////
  if (isImportValid) {
    GLB_Import_arr = new Array(); // to empty the same global array
    var data = SheetImportName.getDataRange().getValues();
    for(i=0;i<data.length;i++){
      // itemNo
      var itemNo = i;

      // date
      var date = new Date(data[i][1]);

      // action
      var act = action_mapping("CTBC", i, data[i][2].toString(), data[i][3].toString(), data[i][4].toString());

      // amount
      var amount = ((act == "TransferOut" || act == "Withdraw") ? data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"") : data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")) * 1;
      // Logger.log("i: %d, act: %s, amount: %d, data: %s: ",i, act, amount, data[i]);

      // balance
      var balance = (data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"")) * 1;
      
      //Logger.log(`CCC: i: ${i}, date: ${date}, act: ${act}, amount: ${amount}, balance: ${balance}, data: ${data[i]}`);
      // from
      var fromNote = note_mapping("CTBC", i, date, balance, act, data[i][6].toString());
      
      if (act != ""){
        GLB_Import_arr.push([date,act,amount,balance,fromNote[0],fromNote[1]]);
      }
    }

    merge_BankRecord("中國信託","000014853**1373*");
  }

  return isParseValid;
}

function load_BankRecord() {
  /**
   * To restore current SheetBankRecordName content
   * [Date,Action,Amount,Balance,FromAccountName,FromAccount,ToAccountName,ToAccount,MergeDate,RecordCheck,RecordNote,ContractOverrid]
   */
  
  const bankRecordRowOfs = 1; // the offset from the top row, A2 is 1
  const bankRecordColOfs = 1; // to exclude itemNo
  const bankRecordContentLen = 12; // include from Date to ContractOverrid
  const bankRecordCompareLen = 6; // include from Date to FromAccount
  
  GLB_BankRecord_arr = new Array();  // Since GLB_BankRecord_arr is a Global array, will regen this content every time merge starting.
  if (SheetBankRecordName.getLastRow()) { // for a non-empty database case
    var data = SheetBankRecordName.getRange(1+bankRecordRowOfs, 1+bankRecordColOfs, SheetBankRecordName.getLastRow()-bankRecordRowOfs, bankRecordContentLen).getValues(); // exclude itemNo column
    for(i=0;i<data.length;i++){
      GLB_BankRecord_arr.push(data[i].concat([null,null,""])); // expand a placeholder for LN_ContractNo, LN_RentProperty, LN_LineMsg
    }

    VAR_CompareBankRecord_arr = new Array();
    var compare_data = SheetBankRecordName.getRange(1+bankRecordRowOfs, 1+bankRecordColOfs, SheetBankRecordName.getLastRow()-bankRecordRowOfs, bankRecordCompareLen).getValues(); // exclude itemNo column and [toAccountName, toAccount, mergeDate, RecordCheck, RecordNote, ContractOverrid, LN_ContractNo,LN_RentProperty]
    for(i=0;i<compare_data.length;i++){
      VAR_CompareBankRecord_arr.push(compare_data[i]);
    }
  }
}

function merge_BankRecord(toAccountName,toAccount) {
  /////////////////////////////////////////
  // Merge existed item
  /////////////////////////////////////////
  for(i=0;i<GLB_Import_arr.length;i++){
    if (VAR_CompareBankRecord_arr.join().indexOf(GLB_Import_arr[i].join()) == -1){
      GLB_BankRecord_arr.push(GLB_Import_arr[i].concat([toAccountName,toAccount,CONST_TODAY_DATE,"","","",null,null,""]));// expand a placeholder for RecordCheck, RecordNote, ContractOverrid, LN_ContractNo, LN_RentProperty, LN_LineMsg
    }
    else {
      // Found duplicated itemRecord, passed.
    }
  }
}

function post_BankRecord() {
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
  // Attach with serial number, GLB_BankRecord_arr item order is correct from here
  /////////////////////////////////////////
  for(var i=0;i<GLB_BankRecord_arr.length;i++){
    GLB_BankRecord_arr[i] = [i].concat(GLB_BankRecord_arr[i]);
  }

  /////////////////////////////////////////
  // Attach with Line Msg
  /////////////////////////////////////////
  append_LineMsg();
}

function write_BankRecord() {
  const bankRecordRowOfs = 1; // the offset from the top row, A2 is 1

  /////////////////////////////////////////
  // Write Out
  /////////////////////////////////////////
  if (VAR_isNewRecordImport) {
    SheetBankRecordName.getRange(1+bankRecordRowOfs,1,GLB_BankRecord_arr.length,GLB_BankRecord_arr[0].length).clearContent();
    SheetBankRecordName.getRange(1+bankRecordRowOfs,1,GLB_BankRecord_arr.length,GLB_BankRecord_arr[0].length).setValues(GLB_BankRecord_arr);
    Logger.log(`[Info] SheetBankRecordName clear width is ${GLB_BankRecord_arr[0].length}, height: ${GLB_BankRecord_arr.length}`);
    Logger.log(`[Info] Size of GLB_BankRecord_arr is ${GLB_BankRecord_arr.length}`);
  } else {
    // only update [RecordCheck	RecordNote	ContractOverrid	LN_ContractNo	LN_RentProperty	LN_LineMsg]
    var itemPos = new itemBankRecord(GLB_BankRecord_arr[0]);
    var partial_BankRecord_arr = new Array();
    for (var i=0;i<GLB_BankRecord_arr.length;i++) {
      var item = new itemBankRecord(GLB_BankRecord_arr[i]);
      var dataArr = new Array();
      for (var j=(item.ColPos_RecordCheck-1);j<item.itemPackMaxLen;j++) {
        dataArr.push([item.itemPack[j]]);
      }
      partial_BankRecord_arr.push(dataArr);
    }

    SheetBankRecordName.getRange(1+bankRecordRowOfs,itemPos.ColPos_RecordCheck,partial_BankRecord_arr.length,partial_BankRecord_arr[0].length).clearContent();
    SheetBankRecordName.getRange(1+bankRecordRowOfs,itemPos.ColPos_RecordCheck,partial_BankRecord_arr.length,partial_BankRecord_arr[0].length).setValues(partial_BankRecord_arr);
    Logger.log(`[Info] SheetBankRecordName clear width is ${partial_BankRecord_arr[0].length}, height: ${partial_BankRecord_arr.length}`);
    Logger.log(`[Info] Size of GLB_BankRecord_arr is ${partial_BankRecord_arr.length}`);
  }
}

function note_mapping(type, id, date, balance, act, note_in) {
  if (act !== "TransferIn" && act !== "Deposit") {
    return ["", ""];
  }

  try {
    if (type === "CTBC") {
      const note_tmp = note_in[6].toString().replace(/\s\*\*\s/g, "**").replace(CONST_This_Account_Number, "");
      const regExp = new RegExp("([0-9]{9}\\*\\*[0-9]{4}[\\*|0-9])", "gi");
      const fromAccount = regExp.exec(note_tmp)[0];
      const fromAccountName = note_tmp.replace(regExp, "").replace(/[\s|\n|\r|\t]/g, "");
      return [fromAccountName, fromAccount];
    } else if (type === "KTB") {
      const fromAccount = note_in[7].toString().replace(/[\s|\n|\r|\t]/g,"");
      const fromAccountName = note_in[6].toString().replace(/[\s|\n|\r|\t]/g,"");
      return [`${fromAccountName};${fromAccount}`, fromAccount];
    } else if (type === "ESUN") {
      const regExp = new RegExp("([0-9]{3}/[0-9]{10}\d*)", "gi");
      const fromAccount = regExp.exec(note_in[2].toString())?.[0] || "";
      const fromAccountName = `${note_in[1].toString().replace(/[\s|\n|\r|\t]/g, "")};${note_in[3].toString().replace(/[\s|\n|\r|\t]/g, "")}`;
      return [fromAccountName, fromAccount];
    } else {
      reportErrMsg(`[note_mapping] import format does not support at id: ${id}??`);
    }
  } catch (e) {
    reportErrMsg(`[note_mapping] An error occurred: ${e.toString()}`);
  }
  return ["", ""];
}

function action_mapping(type, id, act_in, withdraw, deposit) {
  if (withdraw && deposit) console.error("Withdraw and Deposit happened at the same time!!?");
  if (!withdraw && !deposit) console.error("Withdraw and Deposit missed!!?");
  
  const act = act_in.toString().replace(/[\s|\n|\r|\t]/g, "");

  if (type === "CTBC") {
    if (act.match("匯|轉")) {
      return withdraw ? "TransferOut" : "TransferIn";
    }
    if (withdraw && CONST_CTBC_actWithdrawList.includes(act)) return "Withdraw";
    if (deposit && CONST_CTBC_actDepositList.includes(act)) return "Deposit";
  } else if (type === "KTB") {
    if (act.match("轉帳")) {
      return withdraw ? "TransferOut" : "TransferIn";
    }
    if (withdraw && CONST_KTB_actList.includes(act)) return "Withdraw";
    if (deposit && CONST_KTB_actList.includes(act)) return "Deposit";
  } else if (type === "ESUN") {
    if (act.match("匯|轉")) {
      return withdraw ? "TransferOut" : "TransferIn";
    }
    if (act.match("存")) {
      return withdraw ? "Withdraw" : "Deposit";
    }
    return withdraw ? "OtherOut" : "OtherIn";
  } else {
    reportErrMsg(`[action_mapping] import format does not support at id: ${id}??`);
  }
  return "";
}

function rentCollect_parser_chkDuplicated() {
  if (VAR_isNewRecordImport) {
    // less likely duplicated, check for assurance during the new import stage
    if (CFG_Val_obj["CFG_BankRecordCheck_Details"]) {
      for (i=0;i<GLB_BankRecord_arr.length-1;i++){
        for (j=i+1;j<GLB_BankRecord_arr.length;j++){
          var a = new itemBankRecord(GLB_BankRecord_arr[i]);
          if ((a.recordCheck.toString() != "Checked") && (a.amount != 0)) {
            if (a.compare(GLB_BankRecord_arr[j])) {var errMsg = `[rentCollect_parser_chkDuplicated] BankRecord:[${i}]: ${GLB_BankRecord_arr[i]}, BankRecord:[${j}]: ${GLB_BankRecord_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
          }
        }
      }
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
  const tenantStatusColOfs = 3;
  const accountNameColOfs = 4;
  const accountColOfs = 5;
  var all_tenant_chk_arr  = new Array();
  var all_account_chk_arr = new Array();
  SheetTenantName.getRange(1+topRowOfs,1,SheetTenantName.getLastRow()-topRowOfs,1).clear(); // clear itemNo column
  var data = SheetTenantName.getRange(1+topRowOfs,1,SheetTenantName.getLastRow()-topRowOfs,SheetTenantName.getLastColumn()).getValues();
  for(i=0;i<data.length;i++){
    var itemNo = i;
    
    var status = data[i][tenantStatusColOfs].toString().replace(/[\s|\n|\r|\t]/g,"");
    var account_arr = new Array();
    if (data[i].length<accountColOfs+1) {var errMsg = `[rentCollect_parser_Tenant] The tenant account data is missing for ${itemNo}`; reportErrMsg(errMsg);}
    chkNotEmptyEntry(data[i][1]);
    for (var a=accountColOfs;a<data[i].length;a++){
      if (data[i][a].toString().replace(/ /g,'')!='') {
        var account = data[i][a].toString().replace(/ /g,'');
        account_arr.push(account);
        if (new RegExp(".*EndContract.*").exec(status) == null) { all_account_chk_arr.push(account);} // only check those valid contract
        if (chkValidAccount(account) == false) {
          var errMsg = `[rentCollect_parser_Tenant] chkValidAccount failed at itemNo: ${itemNo}`; reportErrMsg(errMsg);
        }
      }
    }

    var tenant = data[i][tenantNameColOfs].toString().replace(/[\s|\n|\r|\t]/g,"");
    var accountName = data[i][accountNameColOfs].toString().replace(/[\s|\n|\r|\t]/g,"");
    all_tenant_chk_arr.push(tenant);
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
  for (i=0;i<all_tenant_chk_arr.length-1;i++){
    for (j=i+1;j<all_tenant_chk_arr.length;j++){
      if (all_tenant_chk_arr[i] == all_tenant_chk_arr[j]) {var errMsg = `[rentCollect_parser_Tenant] tenant: ${all_tenant_chk_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
    }
  }

  // check duplicated account
  for (i=0;i<all_account_chk_arr.length-1;i++){
    for (j=i+1;j<all_account_chk_arr.length;j++){
      var srhMatchPtn = "^" + all_account_chk_arr[i].toString().replace(/[*]/g,"[\u4E00-\uFF5A0-9A-Za-z\u0020-\u007E]?") + "$";
      var srhMatchRegexp = new RegExp(srhMatchPtn,"gi");
      var srhMatchRegexpHit = srhMatchRegexp.exec(all_account_chk_arr[j]);
      if (srhMatchRegexpHit) {var errMsg = `[rentCollect_parser_Tenant] account: ${all_account_chk_arr[j]} is duplicated!`; reportErrMsg(errMsg);}
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

    if (chkValidAccount(data[i][9]) == false) {
      var errMsg = `[rentCollect_parser_Contract] chkValidAccount failed at itemNo: ${itemNo}`; reportErrMsg(errMsg);
    }
    var toAccount = data[i][9].toString().replace(/[\s|\n|\r|\t]/g,"");

    // endContract
    var tmp = data[i][10].toString().replace(/[\s|\n|\r|\t]/g,"");
    var endContract;
    if (new RegExp(".*EndContract.*").exec(tmp) != null) endContract = true
    else endContract = false
    // else {var errMsg = `[rentCollect_parser_Tenant] This is not a valid endContract: ${tmp}`; reportErrMsg(errMsg);}

    // note
    var note = data[i][11].toString().replace(/[\s|\n|\r|\t]/g,"");

    // fileLink
    var fileLink = data[i][12].toString().replace(/[\s|\n|\r|\t]/g,"");

    // proxy
    var proxy = data[i][13].toString().replace(/[\s|\n|\r|\t]/g,"");

    // org
    var org = data[i][14].toString().replace(/[\s|\n|\r|\t]/g,"");

    // endDate
    var chk_illegal = true;
    var endDate = data[i][15];
    
    if ((typeof(endDate) == `string`) && (endDate.replace(/ /g,'') == '')) chk_illegal = false;
    else if ((typeof(endDate) == `object`)) chk_illegal = false;
    else {var errMsg = `[rentCollect_parser_Tenant] This is not a valid endDate: ${endDate}`; reportErrMsg(errMsg);}
    
    GLB_Contract_arr.push([itemNo,fromDate,toDate,rentProperty,deposit,amount,period,tenantName,tenantAccountName_regex,tenantAccount_arr,toAccountName,toAccount,endContract,note,fileLink,proxy,org,endDate]);

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
          GLB_Contract_arr[i][item.ColPos_EndDate+1] = endDate; // pos of endData in GLB_Contract_arr is ColPos_EndDate+1
          SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(endDate); // abnormal end date, terminate contract ahead.
        }
        else if (item.toDate <= CONST_TODAY_DATE) {
          var endDate = new Date(item.toDate.getTime()-CONST_MILLIS_PER_DAY); // toDate is excluded
          GLB_Contract_arr[i][item.ColPos_EndDate+1] = endDate;
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
      GLB_Contract_arr[i][item.ColPos_EndDate+1] = CONST_SuperFeatureDate; // expected to be not used, just for data type consistency
      SheetContractName.getRange(1+topRowOfs+i,item.ColPos_EndDate).setValue(""); 
    }
  
  
  }
  // GLB_Contract_arr.forEach (
  //   function(data) {
  //     var item = new itemContract(data);
  //     Logger.log(`RXX: ${item.show()}, getMaxContractNo: ${getMaxContractNo}`);
  //   }
  // )
  
  // update GLB_Contract_arr relative var
  for(var i=0;i<GLB_Contract_arr.length;i++){
    var item = new itemContract(GLB_Contract_arr[i]);
    VAR_ContractNoList_arr.push(item.itemNo*1);
  }

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
    // 20240917: for regex account, the accout formation is no longer mater
    // var regExpCTBC = new RegExp("([0-9]{9}\\*\\*[0-9]{4}\\*)","gi"); // escape word
    // var meetCTBC = regExpCTBC.exec(account);

    // var regExpKTB = new RegExp("([0-9]{10}\d*)","gi");
    // var meetCKTB = regExpKTB.exec(account);

    // if (meetCTBC || meetCKTB) {
    //   return true;
    // } 
    // else {
    //   if (1) {var errMsg = `[chkValidAccount] This is not a valid account: ${account}`; reportErrMsg(errMsg);}
    //   return false;   
    // }
}

function chkContractIntegrity(){
  // var item = new itemContract(itemIn);
  // Logger.log("AAA: %s",item.contractNo);
  // Logger.log(Object.keys(GLB_Tenant_obj).length);
  // 3. ToDate <= FromDate
  
  ///////////////////////////////////////
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
  
  ///////////////////////////////////////
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
  
  ///////////////////////////////////////
  //ContractOverrid correlate to existed contractNo
  for (i=0;i<GLB_BankRecord_arr.length;i++){ // less likely duplicated, check for assurance
    var record = new itemBankRecord(GLB_BankRecord_arr[i]);
    // Logger.log(`DEBUG20240903: ${record.show()}`)
    if (record.contractOverrid.toString().replace(/[\s|\n|\r|\t]/g,"") != "") {
      if (findContractNoPos(record.contractOverrid)==-1) {
        var errMsg = `[chkContractIntegrity] ContractOverrid contractNo not existed: ${record.itemNo}. ${record.show()}`; reportErrMsg(errMsg);
      } else if (record.recordCheck.toString() != "Checked"){
        var contract = new itemContract(GLB_Contract_arr[findContractNoPos(record.contractOverrid)]);

        var fromDate = new Date(contract.fromDate.getTime()-CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_FromDateMargin"]);
        var toDate   = new Date(contract.toDate.getTime()  +CONST_MILLIS_PER_DAY*CFG_Val_obj["CFG_BankRecordSearch_ToDateMargin"]);
        
        if (!(fromDate<=record.date && record.date<toDate)) {var errMsg = `[chkContractIntegrity] ContractOverrid date not valid: ${record.itemNo}, should be within from ${fromDate} to ${toDate}. \n ${record.show()}. \n ${contract.show()}}`; reportErrMsg(errMsg);}
        
        if (record.toAccountName != contract.toAccountName) {var errMsg = `[chkContractIntegrity] ContractOverrid toAccountName diff: ${record.itemNo}. \n ${record.show()}. \n ${contract.show()}}`; reportErrMsg(errMsg);}

        if (record.toAccount != contract.toAccount) {var errMsg = `[chkContractIntegrity] ContractOverrid toAccount diff: ${record.itemNo}. \n ${record.show()}. \n ${contract.show()}}`; reportErrMsg(errMsg);}
      }
    }
  }
  
  ///////////////////////////////////////
  //Proxying contract must by aligned with proxied contract
  var proxied_list  = new Array();
  // proxying -> proxied
  for(var i=0;i<GLB_Contract_arr.length;i++){
    var contract = new itemContract(GLB_Contract_arr[i]);
    if (contract.isProxying) {
      for (var j=0;j<contract.proxied_array.length;j++) {
        var proxying_contract = contract.proxied_array[j]*1;
        var item = new itemContract(GLB_Contract_arr[findContractNoPos(proxying_contract)]);
        
        // check for no corresponding proxied contract 
        if (item.isProxied == false) {var errMsg = `[chkContractIntegrity] proxying item: ${contract.itemNo}'s proxied item: ${proxying_contract} is not specified as PROXIED!}`; reportErrMsg(errMsg);}
        
        // check for duplicated proxied contract
        if (proxied_list.indexOf(proxying_contract)!=-1) {var errMsg = `[chkContractIntegrity] proxied item: ${proxying_contract} is duplicated. proxied_list: ${proxied_list}!}`; reportErrMsg(errMsg);}

        // for later proxied -> proxying checking
        proxied_list.push(proxying_contract); 
      }
    }
  }

  // proxied -> proxying
  for(var i=0;i<GLB_Contract_arr.length;i++){
    var contract = new itemContract(GLB_Contract_arr[i]);
    if (contract.isProxied) {
      if (proxied_list.indexOf(contract.itemNo)==-1) {var errMsg = `[chkContractIntegrity] proxied item: ${contract.itemNo} is not been proxied by any proxying contract. proxied_list: ${proxied_list}!}`; reportErrMsg(errMsg);}
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

    // contractOverrid
    var contractOverrid = data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    GLB_UtilBill_arr.push([itemNo,date,rentProperty,amount,note,contractOverrid]);
    
    var item = new itemUtilBill(GLB_UtilBill_arr[i]);
    item.show()

    // Write UtilBillNo
    // SheetUtilBillName.getRange(1+topRowOfs+i,1).setValue(itemNo);

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

    // contractOverrid
    var contractOverrid = data[i][6].toString().replace(/[\s|\n|\r|\t]/g,"");
    
    GLB_MiscCost_arr.push([itemNo,date,rentProperty,amount,type,note,contractOverrid]);
    
    var item = new itemMiscCost(GLB_MiscCost_arr[i]);
    item.show()

    // Write UtilBillNo
    // SheetMiscCostName.getRange(1+topRowOfs+i,1).setValue(itemNo);

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

function rentCollect_parser_proxyRecord_MinSheng35() { // 民生路35號報表
  /////////////////////////////////////////
  // README
  // Read MinSheng35 sheet and populate reportWarnGenUtilBill and reportWarnGenMiscCost
  /////////////////////////////////////////
  
  /////////////////////////////////////////
  // Setting
  /////////////////////////////////////////
  var proxyImportNum = CFG_Val_obj["CFG_NewProxyRecordImportNum"].toString().replace(/[\s|\n|\r|\t]/g,"") * 1;
  var numImport = 0;

  // visit sheet of each month
  var allsheets = SheetHandle_MinSheng35.getSheets();
  allsheets.forEach(
    function(sheetName){
      
      // remove all filter
      var filter = sheetName.getFilter();
      if (filter !== null) {
        filter.remove();
      }

      if (sheetName.getName()!="README"){
        if (numImport < proxyImportNum) {
          var warnMsg = `[rentCollect_parser_proxyRecord_MinSheng35] Populate proxy sheet result from @ ${sheetName.getName()}!`;
          reportWarnMsg(warnMsg);
          
          // numImport++
          numImport +=1;

          // apart from README, the reset sheet name align yyyy/mm
          var regex = new RegExp("([0-9]{4}\/[0-9]{2})","gi"); // format must be yyyy/mm
          if (sheetName.getName().toString().match(regex) == null) {
            var errMsg = `[rentCollect_parser_proxyRecord_MinSheng35] incorrect sheetName format: ${sheetName.getName()}.`; reportErrMsg(errMsg);
          }
          
          // parsing 
          var data = sheetName.getDataRange().getValues();
          for(i=0;i<data.length;i++){
            // minSheng35 date structure
            var id          = data[i][0].toString().replace(/[\s|\n|\r|\t]/g,""); // ID row
            var rent        = data[i][1].toString().replace(/[\s|\n|\r|\t]/g,"")*1; //
            var deposit     = data[i][2].toString().replace(/[\s|\n|\r|\t]/g,"")*1;
            var proxy_fee_0 = data[i][3].toString().replace(/[\s|\n|\r|\t]/g,"")*1; // equalt to 1 month rent
            var proxy_fee_1 = data[i][4].toString().replace(/[\s|\n|\r|\t]/g,"")*1; // equalt to 7% month rent
            var util_fee_0  = data[i][5].toString().replace(/[\s|\n|\r|\t]/g,"")*1; // electricity
            var util_fee_1  = data[i][6].toString().replace(/[\s|\n|\r|\t]/g,"")*1; // water
            var other_fee   = data[i][7].toString().replace(/[\s|\n|\r|\t]/g,"")*1; // other fee
            var amount      = data[i][8].toString().replace(/[\s|\n|\r|\t]/g,"")*1; // sum[1:7]
            var note        = data[i][9];
            var fromDate    = data[i][10].toString().replace(/[^\u002E-\u0039]/g,""); // to eliminate everything other that digit or dot

            if ((CONST_MinSheng35_ID.indexOf(id)!=-1) && (fromDate.length>0)){ // only proceed those with ID and fromDate
              
              Logger.log(`[rentCollect_parser_proxyRecord_MinSheng35] ${sheetName.getName()}: ${data[i]}`);

              // Date
              var regex = new RegExp("([0-9]{3}\.[0-9]{2}\.[0-9]{2})","gi"); // format must be yyy.mm.dd
              if (fromDate.match(regex) == null){
                var errMsg = `[rentCollect_parser_proxyRecord_MinSheng35] incorrect fromDate format: ${fromDate}. It should be yyy.mm.dd.`; reportErrMsg(errMsg);
              } else {
                var date = new Date(fromDate);
                date.setSeconds(0);date.setMinutes(0);date.setHours(0);
                date.setFullYear(sheetName.getName().substring(0,4)); // sheetName.yyyy
                date.setMonth(sheetName.getName().substring(5,7)-1); // sheetName.mm, setMonth starts from 0
                date.setDate(fromDate.substring(7,9)); // fromDate.dd
              }
              Logger.log(`[rentCollect_parser_proxyRecord_MinSheng35] date: ${date}`);

              // RentProperty
              var rentProperty = "P_民生路大樓_" + id;

              // Amount
              if (amount != (rent+deposit+proxy_fee_0+proxy_fee_1+util_fee_0+util_fee_1+other_fee)) {
                var errMsg = `[rentCollect_parser_proxyRecord_MinSheng35] incorrect amount @ ${sheetName.getName()}: ID:${id}.`; reportErrMsg(errMsg);
              }

              // populate result
              if (rent!=0){ // for debug only, the rent should come from the BankRecord
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${Math.abs(0)}\t99.Debug\tRent=${rent}, Auto gen debug info:rent @MinSheng sheet:${sheetName.getName()}\n`;
                reportWarnGenMiscCost(warnMsg);
              }
              if (deposit!=0){ // for debug only, the deposit should come from the BankRecord
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${Math.abs(0)}\t99.Debug\tDeposit=${deposit}, Auto gen debug info: deposit @MinSheng sheet:${sheetName.getName()}\n`;
                reportWarnGenMiscCost(warnMsg);
              }
              if (proxy_fee_0!=0){ // equalt to 1 month rent
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${Math.abs(proxy_fee_0)}\t2.Sub_Rent\tAuto gen rent proxy fee @MinSheng sheet:${sheetName.getName()}\n`;
                reportWarnGenMiscCost(warnMsg);
              }
              if (proxy_fee_1!=0){ // equalt to 7% month rent
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${Math.abs(proxy_fee_1)}\t2.Sub_Rent\tAuto gen management proxy fee @MinSheng sheet:${sheetName.getName()}\n`;
                reportWarnGenMiscCost(warnMsg);
              }
              if (other_fee<0){ // already paid in proxy agent side
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${Math.abs(other_fee)}\t2.Sub_Rent\tAuto gen other proxy fee @MinSheng sheet:${sheetName.getName()}\n`;
                reportWarnGenMiscCost(warnMsg);
              }
              if (other_fee>0){ // extra charge fee
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${Math.abs(other_fee)}\t0.Charge_Fee\tAuto gen other proxy fee @MinSheng sheet:${sheetName.getName()}\n`;
                reportWarnGenMiscCost(warnMsg);
              }
              if (util_fee_0!=0){ // electricity
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${util_fee_0}\tAuto gen electricity fee @MinSheng sheet:${sheetName.getName()}, note: ${note}\n`;
                reportWarnGenUtilBill(warnMsg);
              }
              if (util_fee_1!=0){ // water
                var warnMsg = `${Utilities.formatDate(date, 'GMT+8', 'yyyy/MM/dd')}\t${rentProperty}\t${util_fee_1}\tAuto gen water fee @MinSheng sheet:${sheetName.getName()}, note: ${note}\n`;
                reportWarnGenUtilBill(warnMsg);
              }

            }
          }
        }
      }
    }
  )
}

function rentCollect_parser_LineMsg() {
  // for rowPos
  var itemPos = new itemLineMsg();
  var fromDate = CFG_Val_obj["CFG_LineMsg_FromDate"].toString();
  var rowPos=0;
  var dateArr = SheetLineMsgName.getRange(2, 1,  SheetLineMsgName.getLastRow()-1, 3).getValues(); // to fetch [Date, timestamp, userId]
  for (var i=0;i<dateArr.length;i++) {
    var date   = dateArr[i][0];
    var userId = dateArr[i][2];
    if ((rowPos==0) && (new Date(fromDate) <= new Date(date))) {
      rowPos = 2+i; // only update once
      if (1) {var infoMsg = `[rentCollect_parser_LineMsg] LineMsg start parse from row ${rowPos}.`; reportInfoMsg(infoMsg);}
    }
    // if (userId!=CONST_LINE_USERID_ANGENT_0) {var errMsg = `[rentCollect_parser_LineMsg] userID != ${CONST_LINE_USERID_ANGENT_0}}`; reportErrMsg(errMsg);}
  }

  // parse item to array
  var data = SheetLineMsgName.getRange(rowPos, 1, SheetLineMsgName.getLastRow()-rowPos+1, (itemPos.itemPackMaxLen-1)).getValues();
  for(var i=0;i<data.length;i++){
    var dataArr = new Array();
    dataArr.push(i)
    for (var j=0; j < (itemPos.itemPackMaxLen-1); j++) {
      dataArr.push(data[i][j]);
    }
    GLB_LineMsg_arr.push(dataArr);
  }
  
  // to sort array with Date in acending order
  GLB_LineMsg_arr.sort(
    function(x,y){
      var xp = x[1];
      var yp = y[1];
      return xp == yp ? 0 : xp < yp ? -1 : 1;
    }
  )

  // // to show those item
  // for(var i=0;i<GLB_LineMsg_arr.length;i++){
  //   var item = new itemLineMsg();
  //   item.extract(GLB_LineMsg_arr[i])
  //   Logger.log(`item: ${item.show()}`);
  // }
}

function append_LineMsg() {
  var lineMsgRecordNo = 0;
  var CFG_LineMsg_FromDate = Utilities.formatDate(new Date(CFG_Val_obj["CFG_LineMsg_FromDate"].toString()), 'GMT+8', 'yyyy/MM/dd');
  for(var i=0,j=0;i<GLB_LineMsg_arr.length;){
    var item = new itemLineMsg();
    item.extract(GLB_LineMsg_arr[i]);
    // Logger.log(`cur item: ${item.show()}`);

    // merge lineMsg of the within the same date
    var content = item.content;
    do {
      i = ++j; // i start from the next 1st diff
      if (i==GLB_LineMsg_arr.length){
        break;
      } else {
        var next = new itemLineMsg();
        next.extract(GLB_LineMsg_arr[j]);
        // Logger.log(`nxt item: ${next.show()}`);
        if (Utilities.formatDate(item.date, 'GMT+8', 'yyyy/MM/dd') != Utilities.formatDate(next.date, 'GMT+8', 'yyyy/MM/dd')) {
          break;
        } else {
          content += `\n${next.content}`;
        }
      }
    } while (j<GLB_LineMsg_arr.length)
    // Logger.log(`Merged content: ${content}`);

    // Since the date in GLB_BankRecord_arr is in ascending order, ony visit those match date bankRecord
    for (var k=lineMsgRecordNo;k<GLB_BankRecord_arr.length;){
      var record = new itemBankRecord(GLB_BankRecord_arr[k]);
      if (Utilities.formatDate(record.date, 'GMT+8', 'yyyy/MM/dd') < CFG_LineMsg_FromDate) {
        lineMsgRecordNo = ++k;
      } else if (Utilities.formatDate(record.date, 'GMT+8', 'yyyy/MM/dd') <  Utilities.formatDate(item.date, 'GMT+8', 'yyyy/MM/dd')) {
        lineMsgRecordNo = ++k;
      } else if (Utilities.formatDate(record.date, 'GMT+8', 'yyyy/MM/dd') == Utilities.formatDate(item.date, 'GMT+8', 'yyyy/MM/dd')) {
        record.update([content],flag="lineMsg");
        lineMsgRecordNo = ++k;
      } else {
        Logger.log(`record.date: ${Utilities.formatDate(record.date, 'GMT+8', 'yyyy/MM/dd')}, item.date: ${Utilities.formatDate(item.date, 'GMT+8', 'yyyy/MM/dd')}, item: ${item.show()}`)
        if (k==0) {var errMsg = `[append_LineMsg] How comes!?`; reportErrMsg(errMsg);}
        break;
      }
    }
  }
}
