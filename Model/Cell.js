function Cell() {
    this.col;
    this.row;
    this.note;
}
Cell.prototype.indexToNotation = function (row, col) {
    var colName = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z",
    ];
    return colName[col - 1] + row;
};
Cell.prototype.getNotation = function () {
    return (this.notation = this.indexToNotation(this.row, this.col));
};
