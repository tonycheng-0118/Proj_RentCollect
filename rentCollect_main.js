function rentCollect_main() {
  
  //start
  var time_start = new Date();
  Logger.log(`rentCollect_main start @ ${Utilities.formatDate(time_start, "GMT+8", "HH:mm:ss")}`);
  
  // to remove filter from sheet in case of not closing the filter.
  var time_start_filter = new Date();
  removeFilter();
  var time_finish_filter = new Date();
  if (1) {
    var info = `[rentCollect_main] removeFilter filter time_exec(s): ${(time_finish_filter.getTime() - time_start_filter.getTime()) / 1000}`; reportInfoMsg(info);
  }

  // top setting
  var time_start_cfg = new Date();
  mainCFG();
  var time_finish_cfg = new Date();
  if (1) {
    var info = `[rentCollect_main] mainCFG time_exec(s): ${(time_finish_cfg.getTime() - time_start_cfg.getTime()) / 1000}`; reportInfoMsg(info);
  }
  
  // main
  var time_start_parser = new Date();
  rentCollect_parser();
  var time_finish_parser = new Date();
  if (1) {
    var info = `[rentCollect_main] parser time_exec(s): ${(time_finish_parser.getTime() - time_start_parser.getTime()) / 1000}`; reportInfoMsg(info);
  }

  var time_start_contract = new Date();
  rentCollect_contract();
  var time_finish_contract = new Date();
  if (1) {
    var info = `[rentCollect_main] contract time_exec(s): ${(time_finish_contract.getTime() - time_start_contract.getTime()) / 1000}`; reportInfoMsg(info);
  }

  var time_start_report = new Date();
  rentCollect_report();
  var time_finish_report = new Date();
  if (1) {
    var info = `[rentCollect_main] report time_exec(s): ${(time_finish_report.getTime() - time_start_report.getTime()) / 1000}`; reportInfoMsg(info);
  }

  // finish
  var time_exec = (time_finish_report.getTime() - time_start.getTime()) / 1000; // getTime() is ms
  Logger.log(`rentCollect_main finish @ ${Utilities.formatDate(time_finish_report, "GMT+8", "HH:mm:ss")}, time_exec(s): ${time_exec}`);
  if (1) {
    var info = `[rentCollect_main] Overall time_exec(s): ${time_exec}`; reportInfoMsg(info);
  }

  // error log
  rentCollect_debug_print();

}