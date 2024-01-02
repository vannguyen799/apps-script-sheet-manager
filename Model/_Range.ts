export { _Range };

abstract class _Range {
    Sheet: GoogleAppsScript.Spreadsheet.Sheet;
    range: GoogleAppsScript.Spreadsheet.Range;

    get notation(): string {
        return this.getNotation();
    }

    constructor() {}

    getNotation(): string {
        throw new Error("not implement");
    }

    fetch() {
        throw new Error("not implement");
    }

    getValue() {
        return this.getRange().getValue();
    }

    getValues() {
        return this.getRange().getValues();
    }

    getRange() {
        if (!this.range) {
            this.range = this.Sheet.getRange(this.notation);
        }
        return this.range;
    }

    setSheet(Sheet) {
        if (!Sheet) {
            Sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        }
        this.Sheet = Sheet;
        return this;
    }

    setValue(value) {
        this.getRange().setValue(value);
        return this;
    }

    setValues(values) {
        this.getRange().setValues(values);
        return this;
    }

    clearContent() {
        this.getRange().clearContent();
        return this;
    }

    clear() {
        this.getRange().clear();
        return this;
    }
}
