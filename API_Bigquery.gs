function ResultQuerySQL() {
    var Param = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tab');
    var Selec = Param.getRange('B4').getValue();
    var As = Param.getRange('B12').getValue();
    var From = Param.getRange('B14').getValue();
    var Where = Param.getRange('B16').getValue();
    var And_One = Param.getRange('B18').getValue();
    var Between = Param.getRange('B20').getValue();
    var And_Two = Param.getRange('B22').getValue();
    var And_Three = Param.getRange('B24').getValue();
    var Paramenter = Param.getRange('B26').getValue();
    var Query = "SELECT ".concat(Selec, " AS ",
        As, " FROM ",
        From, " WHERE ",
        Where, " AND ",
        And_One, " BETWEEN ",
        "'" + Between + "'", " AND ",
        "'" + And_Two + "'", " AND ",
        And_Three, " ",
        "'" + Paramenter + "'", ')')
        Browser.msgBox(Query)
    return Query
}

function RunBigQuery() {
    var QuerySQL = ResultQuerySQL()

    var projectId = 'projectId';

    var request = {
        query: QuerySQL,
        useLegacySql: false
    };
    Logger.log("Executar");

    try {

        var queryResults = BigQuery.Jobs.query(request, projectId);
        var jobId = queryResults.jobReference.jobId;

        var sleepTimeMs = 500;
        while (!queryResults.jobComplete) {
            Utilities.sleep(sleepTimeMs);
            sleepTimeMs *= 2;
            queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
        }

        // Get all the rows of results.
        var rows = queryResults.rows;
        while (queryResults.pageToken) {
            queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
                pageToken: queryResults.pageToken
            });
            rows = rows.concat(queryResults.rows);
        }

        if (rows) {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tab');
            sheet.clearContents();

            // Append the headers.
            var headers = queryResults.schema.fields.map(function (field) {
                return field.name;
            });
            sheet.appendRow(headers);

            // Append the results.
            var data = new Array(rows.length);
            for (var i = 0; i < rows.length; i++) {
                var cols = rows[i].f;
                data[i] = new Array(cols.length);
                for (var j = 0; j < cols.length; j++) {
                    data[i][j] = cols[j].v;
                }
            }
            sheet.getRange(2, 1, rows.length, headers.length).setValues(data);
         

        } else {
            Logger.log('No rows returned.');
        }
    }
    catch (e) {
        Browser.msgBox("OCORREU O ERROR: " + e)
    }
}
function Clear() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tab');
    sheet.clearContents();
    
}

function Restart() {
    var Param = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parametros');   
    Param.getRange('B4').setValue('Fields');
    Param.getRange('B12').setValue('DATA_CREATION');
    Param.getRange('B14').setValue('datasetId.tableid');
    Param.getRange('B16').setValue('callCenterCode is not null');
    Param.getRange('B18').setValue('creationDate');
    Param.getRange('B2').setValue('2018-09-17');
    Param.getRange('E2').setValue('2018-09-25');
    Param.getRange('B24').setValue('REGEXP_CONTAINS(upper(callCenterCode),');
  
    var SheetsEmails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tab');
    var emails = SheetsEmails.getRange('C2').getValue();

    Param.getRange('B26').setValue(emails);
}


function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('BigQuery')
        .addItem('Executar Query Padrão', 'RunBigQuery')
        .addSeparator()
        .addItem('Restaurar Query Padrão', 'Restart')
        .addSeparator()
        .addItem('Executar Query Personalizada', 'RunBigQueryPersonalite')
        .addSeparator()
        .addItem('Limpar Tabela', 'Clear')
        .addSeparator()
        .addToUi();
};

function RunBigQueryPersonalite() {
    var TableParam = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parametros'); 
    var QueryPersonalite = TableParam.getRange("A35").getValue();
  
    var QuerySQL = "SELECT * FROM datasetId.tableid WHERE email is not null AND creationDate BETWEEN '2018-10-01 00:00:00' AND '2018-11-07 23:59:59' AND REGEXP_CONTAINS(upper(email),'"+
                  QueryPersonalite + "')";


    var projectId = 'projectId';

    var request = {
        query: QuerySQL,
        useLegacySql: false
    };
    Logger.log("Executar");

    try {

        var queryResults = BigQuery.Jobs.query(request, projectId);
        var jobId = queryResults.jobReference.jobId;

        var sleepTimeMs = 500;
        while (!queryResults.jobComplete) {
            Utilities.sleep(sleepTimeMs);
            sleepTimeMs *= 2;
            queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
        }

        // Get all the rows of results.
        var rows = queryResults.rows;
        while (queryResults.pageToken) {
            queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
                pageToken: queryResults.pageToken
            });
            rows = rows.concat(queryResults.rows);
        }

        if (rows) {
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tab');
            sheet.clearContents();

            // Append the headers.
            var headers = queryResults.schema.fields.map(function (field) {
                return field.name;
            });
            sheet.appendRow(headers);

            // Append the results.
            var data = new Array(rows.length);
            for (var i = 0; i < rows.length; i++) {
                var cols = rows[i].f;
                data[i] = new Array(cols.length);
                for (var j = 0; j < cols.length; j++) {
                    data[i][j] = cols[j].v;
                }
            }
            sheet.getRange(2, 1, rows.length, headers.length).setValues(data);

        } else {
            Logger.log('No rows returned.');
        }
      
    }
    catch (e) {
        Browser.msgBox("OCORREU O ERROR: " + e)
    }
}


