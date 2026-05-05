

var SCRIPT_PROPS = PropertiesService.getScriptProperties();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BI Intelligence')
    .addItem('Executive Dashboard', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Business Intelligence Strategist')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}


function runFullBusinessAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();


  var logSheet = ss.getSheetByName("Audit Log") || ss.insertSheet("Audit Log");
  logSheet.clear();
  logSheet.appendRow(["Step", "Status", "Time"]);

  function log(step, status) {
    logSheet.appendRow([step, status, new Date()]);
    SpreadsheetApp.flush();
  }

  log("Start", "Initializing System");

  var sections = ["Analysis", "Trends", "Predictions", "Recovery", "Marketing"];

  sections.forEach((section, index) => {
    log(section, "Generating Data");

    var data = generateDummyData(section);

    writeToTabStructured(ss, section, data);

    Utilities.sleep(800);
    SpreadsheetApp.flush();
  });

  log("Completed", "All Reports Generated");

  return "Done";
}


function generateDummyData(type) {
  var headers = ["Product", "Region", "Month", "Revenue", "Cost", "Profit", "Growth %", "Units Sold", "Category"];

  var rows = [headers];

  for (var i = 1; i <= 14; i++) {
    rows.push([
      "Product " + i,
      ["US","UK","UAE","PK"][i % 4],
      ["Jan","Feb","Mar","Apr"][i % 4],
      Math.floor(Math.random() * 10000 + 2000),
      Math.floor(Math.random() * 5000 + 1000),
      Math.floor(Math.random() * 4000 + 500),
      (Math.random() * 20).toFixed(2) + "%",
      Math.floor(Math.random() * 300 + 50),
      ["A","B","C"][i % 3]
    ]);
  }

  return rows;
}


function writeToTabStructured(ss, tabName, tableData) {
  var sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sheet.clear();

  var range = sheet.getRange(1, 1, tableData.length, tableData[0].length);
  range.setValues(tableData);

  var headerRange = sheet.getRange(1, 1, 1, tableData[0].length);
  headerRange.setBackground("#1e2a38")
    .setFontColor("white")
    .setFontWeight("bold");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, tableData[0].length);


  sheet.activate();
  SpreadsheetApp.flush();
}