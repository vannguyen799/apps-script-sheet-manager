import { Cell } from "./model/Cell";
import { VerticalColumn } from "./model/Column";

export { SheetManager, SheetManagerSingleton };

// @VIEW MODEL
class SheetManager {
    spreadsheetId: string;
    sheetName: string;
    Spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
    Sheet: GoogleAppsScript.Spreadsheet.Sheet;
    data: Array<Array<string>>;
    notes: Array<Array<string>>;

    constructor(sheetName: string, spreadsheetId: string | null = null) {
        this.setSpreadsheet(spreadsheetId).setSheet(sheetName).loadData().loadNote();
    }

    setSpreadsheet(spreadsheetId) {
        this.spreadsheetId = spreadsheetId;
        if (spreadsheetId) this.Spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        else {
            this.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        }

        return this;
    }

    setSheet(sheetName) {
        this.sheetName = sheetName;
        let sheet = this.Spreadsheet.getSheetByName(sheetName);
        if (sheet) this.Sheet = sheet;
        else throw new Error("Couldn't fint this sheet name: got " + sheetName);
        return this;
    }

    loadData() {
        this.data = this.Sheet.getDataRange().getValues();
        return this;
    }

    loadNote() {
        this.notes = this.Sheet.getDataRange().getNotes();
        return this;
    }

    getCellByValue(cellValue: string, fromRow: number = 1): Cell | null {
        let data = this.data;
        for (var i = fromRow - 1; i < data.length; i++) {
            for (var j = 0; j < data[0].length; j++) {
                if (data[i][j] == cellValue) {
                    return new Cell(i + 1, j + 1, cellValue).setSheet(this.Sheet);
                }
            }
        }
        return null;
    }

    _getColumn(headerCell: Cell, numRows = null, ignoreBlank = true): VerticalColumn {
        let data = this.data;
        let cells: Array<Cell> = new Array();

        for (var i = headerCell.row; i < data.length; i++) {
            if (numRows && i - headerCell.row + 1 > numRows) break;
            if (!ignoreBlank) {
                if (data[i][headerCell.col - 1] != "") {
                    cells.push(new Cell(i + 1, headerCell.col, data[i][headerCell.col - 1], null, this.Sheet));
                }
            } else {
                cells.push(new Cell(i + 1, headerCell.col, data[i][headerCell.col - 1], null, this.Sheet));
            }
        }
        return new VerticalColumn(headerCell, cells).setSheet(this.Sheet);
    }

    getColumn(header: string, numRows = null, ignoreBlank = true): VerticalColumn | null {
        let headerCell = this.getCellByValue(header);
        if (headerCell) {
            return this._getColumn(headerCell, numRows, ignoreBlank);
        }
        return null;
    }

    getColumnValues(header: string, numRows = null, ignoreBlank = true): Array<any> | null {
        let headerCell = this.getCellByValue(header);
        if (headerCell) {
            return this._getColumn(headerCell, numRows, ignoreBlank).getData();
        }
        return null;
    }

    setColumnValues(headerCell: Cell, values: string | Array<string>) {
        if (values instanceof String) {
            this.Sheet.getRange(headerCell.row + 1, headerCell.col).setValue(values);
        }
        if (Array.isArray(values)) {
            let range = headerCell.notation[0] + (headerCell.row + 1) + ":" + headerCell.notation[0] + (headerCell.row + 1 + values.length - 1);
            this.Sheet.getRange(range).setValues(VerticalColumn.parseArrayData(values));
        }
        return this;
    }

    clearColumnValues(headerCell: Cell, numRows: number) {
        if (!numRows) {
            var lastRow = headerCell.row + 1;
        } else {
            var lastRow = headerCell.row + 1 + numRows;
        }
        this.Sheet.getRange(headerCell.notation[0] + (headerCell.row + 1) + ":" + headerCell.notation[0] + lastRow).clearContent();
        return this;
    }

    getCellByNote(note: string): Cell | null {
        let notes = this.getSheetNotes();
        for (let i = 0; i < notes.length; i++) {
            if (notes[i].note == note) {
                return notes[i];
            }
        }
        return null;
    }

    getSheetNotes(): Array<Cell> {
        var notes: Array<Cell> = new Array();
        var notesArray = this.notes;
        for (let row = 0; row < notesArray.length; row++) {
            for (let col = 0; col < notesArray[row].length; col++) {
                let note = notesArray[row][col];
                if (note !== "") {
                    let cell = new Cell(row + 1, col + 1, null, note, this.Sheet);
                    notes.push(cell);
                }
            }
        }
        return notes;
    }

    flush() {
        SpreadsheetApp.flush();
        return this;
    }
}

class SheetManagerSingleton extends SheetManager {
    private static instance = {};
    private constructor(sheetName: string, spreadsheetId: string | null = null) {
        if (new.target !== SheetManagerSingleton) {
            throw new TypeError("SheetManagerSingleton is not extendable.");
        }

        super(sheetName, spreadsheetId);
        if (!SheetManagerSingleton.instance[String(spreadsheetId)]) SheetManagerSingleton.instance[String(spreadsheetId)] = {};
        SheetManagerSingleton.instance[String(spreadsheetId)][sheetName] = this;
    }

    getInstance(sheetName: string, spreadsheetId: string | null = null): SheetManagerSingleton {
        if (!SheetManagerSingleton.instance || !SheetManagerSingleton.instance[String(spreadsheetId)] || !SheetManagerSingleton.instance[String(spreadsheetId)][sheetName]) {
            new SheetManagerSingleton(sheetName, spreadsheetId);
        }

        return SheetManagerSingleton.instance[String(spreadsheetId)][sheetName];
    }
}
