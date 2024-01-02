export { GoogleSheetAPI };

class GoogleSheetAPI {
    accessToken: string;
    private static instance: GoogleSheetAPI | null = null;

    private constructor(accessToken = null) {
        this.setAccessToken(accessToken);
        GoogleSheetAPI.instance = this;
    }

    static getInstance() {
        if (!GoogleSheetAPI.instance) {
            GoogleSheetAPI.instance = new GoogleSheetAPI();
        }
        return GoogleSheetAPI.instance;
    }

    setAccessToken(accessToken) {
        this.accessToken = accessToken;
        return this;
    }

    addSheet(spreadsheetId, sheetName) {
        var headers = {
            Authorization: "Bearer " + this.accessToken,
            "Content-Type": "application/json",
        };
        var options = {
            headers: headers,
        };
        var url = "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId + ":batchUpdate";
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

        return response;
    }

    setValues(sheetId, range, values) {
        var headers = {
            Authorization: "Bearer " + this.accessToken,
            "Content-Type": "application/json",
        };
        // The URL for the Sheets API to set values in the range
        var apiUrl = "https://sheets.googleapis.com/v4/spreadsheets/" + sheetId + "/values/" + range + "?valueInputOption=RAW";
        var payload = {
            values: values,
        };
        var response = UrlFetchApp.fetch(apiUrl, {
            method: "put",
            headers: headers,
            payload: JSON.stringify(payload),
        });
        return response;
    }
}
