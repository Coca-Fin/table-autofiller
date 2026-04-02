const SHEET_ID = '';
const SHEET = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
SHEET.getRange("A1:G1").setValues([["Продукт",	"Склад",	"Холодильник", 	"Всего", "В системе",	"Разница", "Цена"]])
const RANGE = `A2:G${SHEET.getLastRow()}`

function doGet() {
  return HtmlService.createHtmlOutputFromFile('form');
}


function formatTable(){
//Форматирование таблицы для удобства восприятия
  const rules = SHEET.getConditionalFormatRules()
  const columnF = SHEET.getRange(`F2:F${SHEET.getLastRow()}`)
  
  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)            // условие
    .setBackground("#d6d500")              // цвет фона
    .setRanges([columnF])                     // диапазон
    .build();
  rules.push(rule)

  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)            // условие
    .setBackground("#059922")              // цвет фона
    .setRanges([columnF])                     // диапазон
    .build();
  rules.push(rule)

  rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)            // условие
    .setBackground("#d61f00")              // цвет фона
    .setRanges([columnF])                     // диапазон
    .build();
  rules.push(rule)

  
  SHEET.setConditionalFormatRules(rules);
  SHEET.getRange(`A1:H${SHEET.getLastRow()}`).setBorder(true, true, true, true, true, true)
}

function setFormuls(){
//Функция заполнение ячеек формулами
  const LASTROW = SHEET.getLastRow()
  SHEET.getRange(`F2:F${LASTROW}`).setFormula("=D2-E2");
  SHEET.getRange(`D2:D${LASTROW}`).setFormula(`=B2+C2`);
  SHEET.getRange(`H2:H${LASTROW}`).setFormula(`=F2*G2`);
  SHEET.getRange(`I${LASTROW}`).setFormula(`=SUM(H2:H${LASTROW})`);
}

function getOldDataAtoEColumns(){
//функция сбора заполненных в таблдице данных
  var oldData = SHEET.getRange(RANGE).getValues()
  return oldData
  
}

function cleanRows(listRows) {
/*Очистка вводных данных
Пример:
str("
напитки
hqd энергетик 0.5л взрывная малина 15 190
сэндвичи
симба 8 195
")
*/
  let i = 0;

  while (i < listRows.length) {
    let isAlphabetic = /^[а-яА-яЁё]+$/.test(listRows[i]);
    // Удалить строки, содержащие только буквы
    if (isAlphabetic) {
      listRows.splice(i, 1);
      continue;
    }

    // Развернуть строку и разбить максимум на 3 части
    
    let parts = listRows[i].split(/\t+/);

    try {
      // Попытка обработать строку
      if (/^\d+$/.test(parts[1])) {
        listRows[i] = [parts[0], parts[1], parts[2]];
        i++;
      } else {
        listRows.splice(i, 1);
      }
    } catch (e) {
      listRows.splice(i, 1);
    }
  }

  // Отсортировать результат
  return listRows.sort();
}

function listToMap(listRows){
//Заполнение Map объекта для работы с данными
  listRows.sort()
  var map = new Map()
  for (let i = 0; i < listRows.length; i++){
  //заполнение элементов строк на основе прошлых данных
    if (listRows[i][4] == null){
      map.set(listRows[i][0], ['', '', '', listRows[i][2], '', listRows[i][1]])
    } else{
      map.set(listRows[i][0], [listRows[i][1], listRows[i][2], listRows[i][3], listRows[i][4], listRows[i][5], listRows[i][6]])
    }
  }
  map.delete('')
  return map
}

function mapSerializer(dataMap){
/*
преобразование Map в список
для использования метода .setValues()
*/
  var listForSave = [];
  var entries = Array.from(dataMap.entries())
  for (let i = 0; i < dataMap.size; i++){
    listForSave.push([entries[i][0], entries[i][1][0], entries[i][1][1], entries[i][1][2], entries[i][1][3], entries[i][1][4], entries[i][1][5]]);
  }
  return listForSave;
}

function compareMap(oldMap, newMap){
  /*
сравнение старых и новых строк для дозаполнения новых строк
на основен уже известных ранее данных.
Решает проблему учета новых товаров,
которых еще не было в таблице
*/
  function compareData(oldMap, newMap, key){
/*
совмещение старых и новых данных
*/
    if (oldMap.has(key) && newMap.has(key)){
      let oldItem = oldMap.get(key);
      let newItem = newMap.get(key);
      return ([key, [oldItem[0], oldItem[1], '', newItem[3], '', newItem[5]]]);
    } else if (newMap.has(key)){
      return [key, Array.from(newMap.get(key))];
    }   
  }
  
  var comparedMap = new Map();

  if (oldMap.size < newMap.size){
    var longArrayKeys = Array.from(newMap.keys());
    var shortArrayKeys = Array.from(oldMap.keys());

  } else {
    var longArrayKeys = Array.from(oldMap.keys());
    var shortArrayKeys = Array.from(newMap.keys());
  }


  for (var i = 0; i < shortArrayKeys.length; i++){
    try{
      if (!comparedMap.has(shortArrayKeys[i])){
        let comparedData = compareData(oldMap, newMap, shortArrayKeys[i])
        comparedMap.set(...comparedData)
      }
    }
    catch(e){}
    try{
      if (!comparedMap.has(longArrayKeys[i])){
        let comparedData = compareData(oldMap, newMap, longArrayKeys[i])
        comparedMap.set(...comparedData)
      }
    }
    catch(e){}
  }
  
  for (i; i < longArrayKeys.length; i++){
    try{
      if (!comparedMap.has(longArrayKeys[i])){
        let comparedData = compareData(oldMap, newMap, longArrayKeys[i])
        if (comparedData){comparedMap.set(...comparedData)}
      }
    }
    catch{
      continue
    }
  }

  return comparedMap
}

function saveData(text){
/*основная функция приложения, 
создающая основную цепочку действий
*/
  var newData = listToMap(cleanRows(text.toLowerCase().split('\n')));
  var oldData = listToMap(getOldDataAtoEColumns());
  if (Array.from(oldData.keys())[0][0] != null){
    var comparedData = compareMap(oldData, newData)
  }else {
    var comparedData = newData
    }
  let valid_list = mapSerializer(comparedData)
  SHEET.getRange(`A2:I${SHEET.getLastRow()+1}`).clear()
  SHEET.getRange(`A2:G${valid_list.length+1}`).setValues(valid_list)
  setFormuls()
  formatTable()
  return [Array.from(newData.entries()), Array.from(oldData.entries())]

}

