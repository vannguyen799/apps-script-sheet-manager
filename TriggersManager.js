function TriggersManager() {
    this.Spreadsheet;
    this.setActiveSpreadsheet();
}
TriggersManager.prototype.setActiveSpreadsheet = function () {
    this.Spreadsheet = SpreadsheetApp.getActive();
};
TriggersManager.prototype.setSpreadsheet = function (spreadsheetId) {
    this.Spreadsheet = SpreadsheetApp.openById(spreadsheetId);
};
TriggersManager.prototype.createTimeDrivenTriggers = function (
    functionName,
    second
) {
    ScriptApp.newTrigger(functionName)
        .timeBased()
        .after(second * 1000)
        .create();
};
TriggersManager.prototype.deleteAllTriggerWithName = function (functionName) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
        if (t.getHandlerFunction() == functionName) ScriptApp.deleteTrigger(t);
    });
};
TriggersManager.prototype.deleteAllTrigger = function () {
    ScriptApp.getProjectTriggers().forEach(function (t) {
        return ScriptApp.deleteTrigger(t);
    });
};
TriggersManager.prototype.createOnOpenTrigger = function (functionName) {
    ScriptApp.newTrigger(functionName)
        .forSpreadsheet(this.Spreadsheet)
        .onOpen()
        .create();
};
TriggersManager.prototype.createOnEditTrigger = function (functionName) {
    ScriptApp.newTrigger(functionName)
        .forSpreadsheet(this.Spreadsheet)
        .onEdit()
        .create();
};
