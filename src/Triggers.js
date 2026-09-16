function removeAllTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
}

function installTriggers() {
  removeAllTriggers_();
  // consolidateCpTasks() calls mirrorCpMaster() itself as its last step, so
  // one daily trigger covers the whole pipeline.
  CONFIG.CP_TRIGGER_TIMES.forEach(function (time) {
    ScriptApp.newTrigger('consolidateCpTasks').timeBased().everyDays(1).atHour(time[0]).nearMinute(time[1]).create();
  });
}
