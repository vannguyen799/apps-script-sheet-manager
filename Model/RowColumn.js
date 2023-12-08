function HorizontalColumn() {
    this.headerCell;
    this.data = [];
}
HorizontalColumn.prototype.valuesArrayToRowArray = function (data) {
    return data.map(function (d) {
        return d.toString();
    });
};
