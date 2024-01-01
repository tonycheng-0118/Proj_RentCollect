function rentCollect_main() {
  
  //start
  var time_start = new Date();
  Logger.log(`rentCollect_main start @ ${Utilities.formatDate(time_start, "GMT+8", "HH:mm:ss")}`);
  
  // to remove filter from sheet in case of not closing the filter.
  removeFilter();

  // top setting
  mainCFG();
  
  // main
  rentCollect_parser();
  rentCollect_contract();
  rentCollect_report();

  // finish
  var time_finish = new Date();
  var time_exec = (time_finish.getTime() - time_start.getTime()) / 1000; // getTime() is ms
  Logger.log(`rentCollect_main finish @ ${Utilities.formatDate(time_finish, "GMT+8", "HH:mm:ss")}, time_exec(s): ${time_exec}`);
  if (1) {
    var info = `[rentCollect_main] time_exec(s): ${time_exec}`; reportInfoMsg(info);
  }

  // error log
  rentCollect_debug_print();

}