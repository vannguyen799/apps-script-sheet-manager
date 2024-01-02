import { Cell } from "./Cell";
import { _Range } from "./_Range";

export { VerticalColumn, HorizontalColumn };

abstract class Column extends _Range {
    headerCell: Cell;
    data: Array<string>;
    cells: Array<Cell>;

    constructor() {
        super();
    }

    setHeaderCell(headerCell: Cell) {
        this.headerCell = headerCell;
        return this;
    }

    setCells(cells: Array<Cell>) {
        this.cells = cells;
        return this;
    }

    getData(index: number | null = null): Array<any> | any {
        let data = new Array();
        for (const cell of this.cells) {
            data.push(cell.value);
        }
        switch (index) {
            case null:
                return data;
            case -1:
                return data[data.length - 1];
            default:
                return data[index];
        }
    }
}

class VerticalColumn extends Column {
    constructor(headerCell: Cell, cells: Array<Cell> = [], Sheet = null) {
        super();
        this.setHeaderCell(headerCell).setCells(cells).setSheet(Sheet);
    }

    static parseArrayData(data: Array<string>): Array<Array<string>> {
        return data.map((d) => [d]);
    }

    getNotation(): string {
        let _ = `${this.headerCell.notation}:${this.headerCell.notation[0]}${this.headerCell.row + this.cells.length}`;
        return _;
    }
}

class HorizontalColumn extends Column {
    constructor(headerCell: Cell, cells: Array<Cell> = [], Sheet = null) {
        super();
        this.setHeaderCell(headerCell).setCells(cells).setSheet(Sheet);
    }

    static parseArrayData(values: Array<string>): Array<string> {
        return values.map(function (d) {
            return d.toString();
        });
    }
}
