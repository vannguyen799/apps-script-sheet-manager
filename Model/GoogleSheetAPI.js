function GoogleSheetAPI() {
    this.accessToken;
}
GoogleSheetAPI.prototype.addSheet = function (spreadsheetId, sheetName) {
    var headers = {
        Authorization: "Bearer " + this.accessToken,
        "Content-Type": "application/json",
    };
    var options = {
        headers: headers,
    };
    var url =
        "https://sheets.googleapis.com/v4/spreadsheets/" +
        spreadsheetId +
        ":batchUpdate";
    var payload = {
        requests: [
            {
                addSheet: {
                    properties: {
                        title: sheetName,
                        gridProperties: {
                            rowCount: 20,
                            columnCount: 6,
                        },
                    },
                },
            },
        ],
    };
    var response = UrlFetchApp.fetch(url, {
        method: "post",
        headers: options.headers,
        contentType: "application/json",
        payload: JSON.stringify(payload),
    });
};
GoogleSheetAPI.prototype.setValues = function (sheetId, range, values) {
    var headers = {
        Authorization: "Bearer " + this.accessToken,
        "Content-Type": "application/json",
    };
    // The URL for the Sheets API to set values in the range
    var apiUrl =
        "https://sheets.googleapis.com/v4/spreadsheets/" +
        sheetId +
        "/values/" +
        range +
        "?valueInputOption=RAW";
    var payload = {
        values: values,
    };
    var response = UrlFetchApp.fetch(apiUrl, {
        method: "put",
        headers: headers,
        payload: JSON.stringify(payload),
    });
};
