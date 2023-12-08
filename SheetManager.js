// @VIEW MODEL
function SheetManager(sheetName) {
    if (!sheetName) return new Error("Please provide sheet name");
    this.setSheet(sheetName);
    this.loadData();
}

SheetManager.prototype.setSheet = function (sheetName) {
    this.sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
};

SheetManager.prototype.loadData = function () {
    return (this.data = this.sheet.getDataRange().getValues());
};

SheetManager.prototype.getCellByValue = function (cellValue, fromRow = 1) {
    var cell = new Cell();
    var data = this.data;
    for (var i = fromRow - 1; i < data.length; i++) {
        for (var j = 0; j < data[0].length; j++) {
            if (data[i][j] == cellValue) {
                cell.col = j + 1;
                cell.row = i + 1;
                cell.notation = cell.indexToNotation(i + 1, j + 1);
                cell.value = cellValue;
            }
        }
    }
    return cell;
};

SheetManager.prototype.getColumnValues = function (
    headerCell,
    numRows = 100000,
    ignoreBlank = true
) {
    var data = this.data;
    var column = new Column();
    column.headerCell = headerCell;

    for (var i = headerCell.row; i < data.length; i++) {
        if (i - headerCell.row + 1 > numRows) break;
        if (!ignoreBlank) {
            if (data[i][headerCell.col - 1] != "") {
                column.data.push(data[+i][headerCell.col - 1]);
            }
        } else {
            column.data.push(data[+i][headerCell.col - 1]);
        }
    }
    return column;
};

SheetManager.prototype.setColumnValues = function (headerCell, values) {
    if (values instanceof String) {
        this.sheet
            .getRange(headerCell.row + 1, headerCell.col)
            .setValue(values);
    }
    if (Array.isArray(values)) {
        let range =
            headerCell.notation[0] +
            (headerCell.row + 1) +
            ":" +
            headerCell.notation[0] +
            (headerCell.row + 1 + values.length - 1);
        this.sheet
            .getRange(range)
            .setValues(new Column().valuesArrayToColumnArray(values));
    }
};

SheetManager.prototype.clearColumnValues = function (headerCell, numRows) {
    if (!numRows) {
        var lastRow = headerCell.row + 1;
    } else {
        var lastRow = headerCell.row + 1 + numRows;
    }
    this.sheet
        .getRange(
            headerCell.notation[0] +
                (headerCell.row + 1) +
                ":" +
                headerCell.notation[0] +
                lastRow
        )
        .clearContent();
};
SheetManager.prototype.getCellByNote = function (note) {
    var notes = this.getSheetNotes();
    for (var i = 0; i < notes.length; i++) {
        if (notes[i].note == note) {
            return notes[i];
        }
    }
};

SheetManager.prototype.getSheetNotes = function () {
    var notes = new Array();
    var dataRange = this.sheet.getDataRange();
    var notesArray = dataRange.getNotes();
    for (var row = 0; row < notesArray.length; row++) {
        for (var col = 0; col < notesArray[row].length; col++) {
            var note = notesArray[row][col];
            if (note !== "") {
                var cell = new Cell();
                cell.row = row + 1;
                cell.col = col + 1;
                cell.note = note;
                cell.getNotation();
                notes.push(cell);
            }
        }
    }
    return notes;
};

function newSheetManager(sheetName) {
    return new SheetManager(sheetName);
}
