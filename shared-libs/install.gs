// shared-libs/install.gs
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('Troop Automation')
    .addItem('Install/Repair Triggers', 'installTriggers')
    .addToUi();
}
function installTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runDailyDigest').timeBased().atHour(17).everyDays(1).create();
  ScriptApp.newTrigger('hourlyMaintenance').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('Triggers installed.');
}
