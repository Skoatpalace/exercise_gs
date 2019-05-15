function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Advanced')
  .addItem('content', 'buildContent')
  .addItem('Formula', 'addFormula')
  .addSeparator()
  .addItem('Info', 'showInfo')
  .addItem('Dashboard', 'showDashboard')
  .addToUi();
}

function findSheetData(sheetNameSel){
  var ss = SpreadsheetApp.openById('1kBgTMLpmCgkltBuuDJDA1F0p-G8-OY-K96621FU4LOQ');
  var sheet = ss.getSheetByName(sheetNameSel);
  if(sheet != null){
    var data = sheet.getDataRange().getValues();
    Logger.log(data);
    return {'success':true, 'sheetNameSel':sheetNameSel , 'data':data};
  }
  return {'success':false, 'sheetNameSel':sheetNameSel};
}

function showDashboard(){
  var ss = SpreadsheetApp.openById('1kBgTMLpmCgkltBuuDJDA1F0p-G8-OY-K96621FU4LOQ');
  var sheets = ss.getSheets();
  var sheetName;
  var sheetData = [];
  for(var i=0;i<sheets.length;i++){
    sheetName = sheets[i].getName();
    sheetData.push(sheetName);
  }
  var t = HtmlService.createTemplateFromFile('dashboard');
  t.data = {sheets:sheetData};
  var html = t.evaluate().setWidth(1200).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Project Exercise');
}

function buildContent() {
  var ss = SpreadsheetApp.openById('1kBgTMLpmCgkltBuuDJDA1F0p-G8-OY-K96621FU4LOQ');
  var sheet;
  var valueToCopy = ss.getSheetByName('main').getRange(1, 1, 10, 1).getValues(); 
  for(var x=1; x<5;x++){
    sheet = ss.getSheetByName('New Sheet' + x);
    if(sheet != null){
      ss.deleteSheet(sheet);
    }
    sheet = ss.insertSheet();
    sheet.setName('New Sheet' + x);
    for(var col=1; col<4; col++){
      for(var row=1; row<11; row++){
        var randNum = Math.ceil(Math.random()*1000);
        sheet.getRange(row, col).setValue(randNum);        
      }
    }
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1, 10, 1).setValues(valueToCopy);
  }
  Logger.log('build');
}

function showInfo() {
  Logger.log('info');
  var t = HtmlService.createTemplateFromFile('info');
  var html = t.evaluate().setWidth(1200).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Project Exercise');
}

function addFormula(){
  var ss = SpreadsheetApp.openById('1kBgTMLpmCgkltBuuDJDA1F0p-G8-OY-K96621FU4LOQ');
  var sheet = ss.getSheetByName('New Sheet1');
  var range = sheet.getRange('E1:E10');
  range.setFormula('B1+C1+D1');
  range.setFontColor('red');
  range.setFontWeight('bold');
}
