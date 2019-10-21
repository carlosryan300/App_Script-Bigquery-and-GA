function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('BigQuery')
        .addSeparator()
        .addItem('Executar Query', 'Promise')
        .addSeparator()
        .addToUi();
};

function Data_Atual(){
  var now = new Date();
  var noonString = Utilities.formatDate(now, "GTM",'yyyMMdd')
  return noonString
}

function Promise(){
  var TableParam = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Query"); 
  var Data_ = Data_Atual()
  var lista = {
    "Consul":TableParam.getRange("A3").getValue() + Data_ + TableParam.getRange("A4").getValue(),
    "CompraCerta": TableParam.getRange("A7").getValue() + Data_ + TableParam.getRange("A8").getValue(),
    "Brastemp": TableParam.getRange("A11").getValue() + Data_ + TableParam.getRange("A12").getValue()
  }  
  try{
    for (var i in lista){
      //Browser.msgBox(lista[i])
       RunQuery(lista[i])
  }
  }catch (e) {
          Browser.msgBox("OCORREU O ERROR: " + e)
      }

}

function RunQuery(Query) { 
    var projectId = 'whirlpool-lar-datalake';
    var request = {
        query: Query,
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
            var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Total');            
            
            var headers = queryResults.schema.fields.map(function (field) {
                return field.name;
            });
            //sheet.appendRow(headers);
            
            // Append the results.
            var data = new Array(rows.length);
            for (var i = 0; i < rows.length; i++) {
                var cols = rows[i].f;
                data[i] = new Array(cols.length);
                for (var j = 0; j < cols.length; j++) {
                    data[i][j] = cols[j].v;
                }
            }
            sheet.getRange(sheet.getLastRow()+1, 1, rows.length, headers.length).setValues(data);
        } else {
            Logger.log('No rows returned.');
        }
      
    }
    catch (e) {
        Browser.msgBox("OCORREU O ERROR: " + e)
    }
}

function Clear() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Total');
    sheet.clearContents();
    headers = [
        "Loja",	
        "Data",
        "Hora_Particionada",
        "Sessoes",
        "Transacoes",
        "Ultima_Hora_UTC",
        "Ultima_Hora_BR"
    ]
    sheet.appendRow(headers)
    
    
}

