var ss = SpreadsheetApp.getActiveSpreadsheet();


function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu("Отчет")
  .addItem("Отчет по кол-ву квартир", "otchet_kvartir")
  .addItem("Отчет по обходам", "otchet_obhod")
  .addToUi();
}

function findTechnik(name){
  var summa = [];
  var mesto = [];
  var otvet = [];
  
  for (var k = 0; k < 3; k++) {
  var adress = ss.getSheets()[k].getRange(2, 1, ss.getSheets()[k].getLastRow()-1).getValues();
  var data = ss.getSheets()[k].getRange(2, 3, ss.getSheets()[k].getLastRow()-1).getValues();
    
  for (var i = 0; i < data.length; i++) {
    var start = data[i][0].indexOf('#' + name);
    var end = data[i][0].indexOf('#', start + 1);
    if (end == -1) {
      var result = data[i][0].slice(start).trim();
      var arr = result.split(':');
      
      if (arr[1] != undefined) {
        var kol = arr[1].split(',').length;
        summa.push(kol);
        mesto.push(`${arr[0]}:${arr[1]} (${kol}шт), ${adress[i][0]}\n`);
      }
      
    } else if (end != -1) {
      var result = data[i][0].slice(start, end).trim();
      var arr = result.split(':');
      
      if (arr[1] != undefined) {
        var kol = arr[1].split(',').length;
        summa.push(kol);
        mesto.push(`${arr[0]}:${arr[1]} (${kol}шт), ${adress[i][0]}\n`);
      }
      
    }
    
  }
  }
  
  var vsego = summa.reduce((x, y) => x + y, 0);
  otvet.push(mesto);
  otvet.push([].concat(vsego));
  
  return otvet
  
}


function otchet_kvartir() {
  var sheet = ss.getSheetByName('кв-ры по тех');
  var names = sheet.getRange(2, 1, sheet.getLastRow()-1).getValues().filter(item => item != '');
  
  
  names.forEach((item, index) => {
                var result = findTechnik(item[0]);
                sheet.getRange(index + 2, 2).setValue(result[1]);
                sheet.getRange(index + 2, 3).setValue(result[0].toString().trim());
                })           
  
}

function otchet_obhod() {
  var sheet = ss.getSheetByName('обходы');
  var names = sheet.getRange(2, 1, sheet.getLastRow()-1).getValues().filter(item => item != '');
  
  
  names.forEach((item, index) => {
                var result = findTechnik(item[0]);
                sheet.getRange(index + 2, 2).setValue(result[1]);
                sheet.getRange(index + 2, 3).setValue(result[0].toString().trim());
                })           
  
}




