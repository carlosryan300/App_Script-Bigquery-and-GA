function listAccounts() {
  try {
    var accounts = Analytics.Management.Accounts.list();
    var ListName = ''
    var ListId = []
    if (accounts.items && accounts.items.length) {
      for (var i = 0; i < accounts.items.length; i++) {
        var account = accounts.items[i];
        ListName = (account.name);
        ListId.push(account.id);
      }
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameter');
      spreadsheet.getRange('B2').setValue(ListName);
      spreadsheet.getRange('C2').setValue(ListId);
      listWebProperties(ListId);
    } else {
      Browser.msgBox(Logger.log('No accounts found.'));
    }
  } catch (e) {
    Browser.msgBox(e)
  }
};

function listWebProperties(accountId) {
  var ListProp = []
  var ListId = []
  try {
    var webProperties = Analytics.Management.Webproperties.list(accountId);
    if (webProperties.items && webProperties.items.length) {
      for (var i = 0; i < webProperties.items.length; i++) {
        var webProperty = webProperties.items[i];
        ListProp.push('Name:[' + webProperty.name + ']  ID:[' + webProperty.id + ']');
        ListId.push(webProperty.id);
      }
      CreateList('B3', ListProp)
      Segments();
    } else {
      Logger.log('\tNo web properties found.');
    }
  } catch (e) {
    Browser.msgBox(e)
  }
};
function Segments() {
  var segments = Analytics.Management.Segments.list().items
  var ListSegment = []
  var ListMetrics = []
  try {
    if (segments) {
      for (var i = 0, segment; segment = segments[i]; i++) {
        ListSegment.push('Name:[' + segment.name + '] ' + ' ID:[' + segment.segmentId + ']')
      }
      CreateList('B7', ListSegment)
    }
  } catch (e) {
    Browser.msgBox(e)
  }

};

function ListView() {
  var ListPro = []
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameter');
  try {
    var profiles = Analytics.Management.Profiles.list(spreadsheet.getRange("C2").getValue(), spreadsheet.getRange("C3").getValue());

    if (profiles.items && profiles.items.length) {
      for (var i = 0; i < profiles.items.length; i++) {
        var profile = profiles.items[i];
        ListPro.push('Name:[' + profile.name + ']  ID:[' + profile.id + ']')
      }
      CreateList('B4', ListPro)
      Segments();
      //
    } else {
      Logger.log('\t\tNo web properties found.');
    }
  } catch (e) {
    Browser.msgBox(e)
  }

};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Google Analitycs')
    .addItem('Atualizar Property', 'listAccounts')
    .addSeparator()
    .addItem('Atualizar View', 'ListView')
    .addSeparator()
    .addItem('Atualizar Dados', 'ResultAnalytics')
    .addToUi();
};

function ResultAnalytics() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameter');
  var startDate = spreadsheet.getRange('C8').getValue();
  var endDate = spreadsheet.getRange('C9').getValue();

  var tableId = 'ga:' + spreadsheet.getRange('C4').getValue();
  var metric = spreadsheet.getRange('B10').getValue();
  var dimensions = spreadsheet.getRange('b11').getValue();
  var segmentId = spreadsheet.getRange('C7').getValue();

  var options = {
    'dimensions': dimensions,
    'max-results': 10000,
    'segmentId': segmentId
  }
  try {
    var report = Analytics.Data.Ga.get(tableId, startDate, endDate, metric, options);
    fillSheet(report);
    Browser.msgBox(report);
  } catch (e) {
    Browser.msgBox(e)
  }

};
function fillSheet(report) {
  var spread = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameter');
  var nameTable = spread.getRange('B14').getValue();
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameTable);
    spreadsheet.clearContents();

    if (report.rows) {

      // Append the headers.
      var headers = report.columnHeaders.map(function (columnHeader) {
        return columnHeader.name;
      });
      spreadsheet.appendRow(headers);

      // Append the results.
      spreadsheet.getRange(2, 1, report.rows.length, headers.length)
        .setValues(report.rows);

    } else {
      Logger.log('No rows returned.');
    }
  } catch (e) {
    Browser.msgBox(e)
  }
};


function CreateList(Cell, List) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('parameter');
  spreadsheet.getRange(Cell).setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(true)
    .requireValueInList(List, true)
    .build());
};

