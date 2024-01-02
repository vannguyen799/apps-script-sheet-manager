export { TriggersManager };

class TriggersManager {
    spreadsheetId: string;
    Spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
    Sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private static instance = {};

    private constructor(spreadsheetId = null) {
        this.setSpreadsheet(spreadsheetId);

        TriggersManager.instance[String(spreadsheetId)] = this;
    }

    static getInstance(spreadsheetId = null): TriggersManager {
        if (!TriggersManager.instance[String(spreadsheetId)]) {
            new TriggersManager(spreadsheetId);
        }
        return TriggersManager.instance[String(spreadsheetId)];
    }

    setActiveSpreadsheet() {
        this.Spreadsheet = SpreadsheetApp.getActive();
    }

    setSpreadsheet(spreadsheetId: null | string) {
        if (spreadsheetId) this.Spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        else this.setActiveSpreadsheet();
    }

    createTimeDrivenTriggers(functionName, second) {
        ScriptApp.newTrigger(functionName)
            .timeBased()
            .after(second * 1000)
            .create();
    }

    deleteAllTriggerWithName(functionName) {
        ScriptApp.getProjectTriggers().forEach(function (t) {
            if (t.getHandlerFunction() == functionName) ScriptApp.deleteTrigger(t);
        });
    }

    createOnOpenTrigger(functionName) {
        ScriptApp.newTrigger(functionName).forSpreadsheet(this.Spreadsheet).onOpen().create();
    }

    createOnEditTrigger(functionName) {
        ScriptApp.newTrigger(functionName).forSpreadsheet(this.Spreadsheet).onEdit().create();
    }
}
