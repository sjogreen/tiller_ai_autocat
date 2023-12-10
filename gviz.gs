(function(context) {

    const Utils = (context.Utils || (context.Utils = {}));


    /**
     * Queries a spreadsheet using Google Visualization API's Datasoure Url.
     *
     * @param        {String} ssId    Spreadsheet ID.
     * @param        {String} query   Query string.
     * @param {String|Number} sheetId Sheet Id (gid if number, name if string). [OPTIONAL]
     * @param        {String} range   Range                                     [OPTIONAL]
     * @param        {Number} headers Header rows.                              [OPTIONAL]
     */
    Utils.gvizQuery = function(ssId, query, sheetId, range, headers) {
        var response = JSON.parse( UrlFetchApp
                .fetch(
                    Utilities.formatString(
                        "https://docs.google.com/spreadsheets/d/%s/gviz/tq?tq=%s%s%s%s",
                        ssId,
                        encodeURIComponent(query),
                        (typeof sheetId === "number") ? "&gid=" + sheetId :
                        (typeof sheetId === "string") ? "&sheet=" + sheetId :
                        "",
                        (typeof range === "string") ? "&range=" + range :
                        "",
                        "&headers=" + ((typeof headers === "number" && headers > 0) ? headers : "0")
                    ), 
                    {
                        "headers":{
                            "Authorization":"Bearer " + ScriptApp.getOAuthToken()
                        }
                    }
                )
                .getContentText()
                .replace("/*O_o*/\n", "") // remove JSONP wrapper
                .replace(/(google\.visualization\.Query\.setResponse\()|(\);)/gm, "") // remove JSONP wrapper
            ),
            table = response.table,
            rows;

        if (typeof headers === "number") {
            rows = table.rows.map(function(row) {
                return table.cols.reduce(
                    function(acc, col, colIndex) {
                        acc[col.label] = row.c[colIndex] && row.c[colIndex].v;
                        return acc;
                    }, 
                    {}
                );
            });

        } else {
            rows = table.rows.map(function(row) {
                return row.c.reduce(
                    function(acc, col) {
                        acc.push(col && col.v);
                        return acc;
                    },
                    []
                );
            });
        }

        return rows;

    };

    Object.freeze(Utils);

})(this);
