function otchet_kvartir() {
  var sheet = ss.getSheetByName('кв-ры по тех');
  var names = sheet.getRange(2, 1, sheet.getLastRow()-1).getValues().filter(item => item != '');
  
  
  names.forEach((item, index) => {
                var result = findTechnik(item[0]);
                sheet.getRange(index + 2, 2).setValue(result[1]);
                sheet.getRange(index + 2, 3).setValue(result[0].toString().trim());
                })           
  
}
