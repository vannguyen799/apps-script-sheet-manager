import { _Range } from "./_Range";

export { Cell };

class Cell extends _Range {
    col: number;
    row: number;
    value: string | null;
    note: string | null;
    Sheet: GoogleAppsScript.Spreadsheet.Sheet;

    constructor(row: number, col: number, value: string | null = null, note: string | null = null, Sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null) {
        super();
        this.setSheet(Sheet).setRow(row).setColumn(col);
        if (value) this.value = value;
        else this.value = this.getValue();

        this.note = note;
    }

    setRow(row: number) {
        this.row = row;
        return this;
    }

    setColumn(col: number) {
        this.col = col;
        return this;
    }

    getNotation() {
        return Cell.indexToNotation(this.row, this.col);
    }

    static indexToNotation = function (row, col) {
        let colName = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
        return colName[col - 1] + row;
    };
}
