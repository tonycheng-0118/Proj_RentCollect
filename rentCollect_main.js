function rentCollect_main() {
  // Start the timer and log the start time
  const startTime = new Date();
  Logger.log(`rentCollect_main start @ ${Utilities.formatDate(startTime, "GMT+8", "HH:mm:ss")}`);

  // Execute and time the removeFilter function
  let taskStartTime = new Date();
  removeFilter();
  let taskEndTime = new Date();
  reportInfoMsg(`[rentCollect_main] removeFilter time_exec(s): ${(taskEndTime.getTime() - taskStartTime.getTime()) / 1000}`);

  // Execute and time the mainCFG function
  taskStartTime = new Date();
  mainCFG();
  taskEndTime = new Date();
  reportInfoMsg(`[rentCollect_main] mainCFG time_exec(s): ${(taskEndTime.getTime() - taskStartTime.getTime()) / 1000}`);

  // Execute and time the rentCollect_parser function
  taskStartTime = new Date();
  rentCollect_parser();
  taskEndTime = new Date();
  reportInfoMsg(`[rentCollect_main] parser time_exec(s): ${(taskEndTime.getTime() - taskStartTime.getTime()) / 1000}`);

  // Execute and time the rentCollect_contract function
  taskStartTime = new Date();
  rentCollect_contract();
  taskEndTime = new Date();
  reportInfoMsg(`[rentCollect_main] contract time_exec(s): ${(taskEndTime.getTime() - taskStartTime.getTime()) / 1000}`);

  // Execute and time the rentCollect_report function
  taskStartTime = new Date();
  rentCollect_report();
  taskEndTime = new Date();
  reportInfoMsg(`[rentCollect_main] report time_exec(s): ${(taskEndTime.getTime() - taskStartTime.getTime()) / 1000}`);

  // Log the final execution time
  const totalExecutionTime = (taskEndTime.getTime() - startTime.getTime()) / 1000;
  Logger.log(`rentCollect_main finish @ ${Utilities.formatDate(taskEndTime, "GMT+8", "HH:mm:ss")}, time_exec(s): ${totalExecutionTime}`);
  reportInfoMsg(`[rentCollect_main] Overall time_exec(s): ${totalExecutionTime}`);

  // Log any errors
  taskStartTime = new Date();
  rentCollect_debug_print();
  taskEndTime = new Date();
  Logger.log(`[rentCollect_main] debug_print time_exec(s): ${(taskEndTime.getTime() - taskStartTime.getTime()) / 1000}`);
}