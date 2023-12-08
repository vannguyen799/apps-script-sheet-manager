function Column() {
    this.headerCell;
    this.data = [];
}
Column.prototype.valuesArrayToColumnArray = function (data) {
    return data.map((d) => [d]);
};
