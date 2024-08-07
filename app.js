var sheet = SpreadsheetApp.getActiveSpreadsheet();
var crawlingData_sheet = sheet.getSheetByName("crawling_data");
var history_sheet_10min = sheet.getSheetByName("history_10min");
var history_sheet_day = sheet.getSheetByName("history_1day");
var accumulate_sheet = sheet.getSheetByName("누적");
var workerInfo = crawlingData_sheet.getRange("a2:d50").getValues().filter((item)=>{
  return item[0] !== '';
})

function upDateTime() {
  var date = Utilities.formatDate(new Date(), "GMT+9:00", "yyyy-MM-dd'/'HH:mm").split("/");

  return date;
}

function workeStemp10min() {
  var values = getWorkerStatistics();
  var valuesRowLength = values.length;
  var valuesColLength = values[0].length;
  var history_sheet_10min_startLine = history_sheet_10min.getRange("a1").getValue();
  var accumulate_sheet_startLine = accumulate_sheet.getRange("a1").getValue();
  var accumulate = accumulate_sheet.getRange(accumulate_sheet_startLine,2,valuesRowLength,valuesColLength).getValues();

  for(let i=0; i < valuesRowLength; i++) {
    for(let j=0; j < valuesColLength; j++) {
      if(j < 3) {
        accumulate[i][j] = values[i][j];
      } else {
        accumulate[i][j] = Number(values[i][j]) - Number(accumulate[i][j]);
      }
    }
  }

  accumulate_sheet.getRange(accumulate_sheet_startLine,2,valuesRowLength,valuesColLength).setValues(values);
  history_sheet_10min.getRange(history_sheet_10min_startLine,2,valuesRowLength,valuesColLength).setValues(accumulate);
}

function workeStempDay() {
  var values = getWorkerStatistics();
  var valuesRowLength = values.length;
  var valuesColLength = values[0].length;
  var history_sheet_day_startLine = sheet.getSheetByName("history_day").getRange("a1").getValue();
  var accumulate_sheet_startLine = accumulate_sheet.getRange("a1").getValue();
  var accumulate = accumulate_sheet.getRange(accumulate_sheet_startLine,2,valuesRowLength,valuesColLength).getValues();
  
  history_sheet_10min.getRange(history_sheet_day_startLine,2,valuesRowLength,valuesColLength).setValues(values);
}

function getSymbolAmountUsed() {
  var test = sheet.getSheetByName("전체 심볼 사용량");
  var statistics = [["심볼명","코코 심볼 개수","마스 심볼 개수","종합"]];
  var getData1 = [];
  var copyGetData = [];
  var workerIdList1 = getWorkerId(1);

  for(let i=0; i<workerIdList1.length; i++) {
    var copyDataInfo1 = {
      writer: workerIdList1[i][0],
      address: workerIdList1[i][1],
      sheetName:'데이터',
      range:'a11:c47',
    }
    if (copyDataInfo1.address !== '-') {
      getData1[i] = getSheetData(copyDataInfo1);
      for(let j=0; j<getData1[i].length; j++) {
        getData1[i][j][3] = copyDataInfo1.writer;
      }
    } else {
      var emptyData = {
        writer: workerIdList1[i][0],
        address: workerIdList1[0][1],
        sheetName:'데이터',
        range:'a11:c47',
      }
      getData1[i] = getSheetData(emptyData);
      for(let j=0; j<getData1[i].length; j++) {
        getData1[i][j][3] = copyDataInfo1.writer;
      }
    }
  }

  for(let i=0; i < getData1.length; i++) {
    for(let j=0; j < getData1[i].length; j++) {
      Logger.log(getData1[i][j])
      copyGetData.push(getData1[i][j]);
    }
  }

  for(let i=0; i < getData1.length; i++) {
    if (i === 0) {
      statistics = [statistics[0], ...getData1[0]];
    } else {
      for(let j=0; j < statistics.length-1; j++) {
        statistics[j+1][0] = getData1[i][j][0];
        statistics[j+1][1] += getData1[i][j][1];
        statistics[j+1][2] += getData1[i][j][2];
        statistics[j+1][3] = '종합';
      }
    }
  }
  Logger.log(copyGetData)
  Logger.log(copyGetData.length)
  Logger.log(statistics)
  Logger.log(statistics.length)
  statistics = [statistics[0],...copyGetData];
  Logger.log(statistics)
  Logger.log(statistics.length)
  test.getRange(`A1:D${statistics.length}`).setValues(statistics);
}

function getWorkerStatistics() {
  var statistics = [];
  var getData1 = [];
  var getData2 = [];
  var getData3 = [];
  var workerIdList1 = getWorkerId(1);
  var workerIdList2 = getWorkerId(2);

  for(let i=1; i<workerIdList1.length; i++) {
    var copyDataInfo1 = {
      writer: workerIdList1[i][0],
      address: workerIdList1[i][1],
      sheetName:'데이터',
      range:'a2:l2',
    }
    if (copyDataInfo1.address !== '-') {
      getData1[i] = getSheetData(copyDataInfo1);
    } else {
      var emptyData = {
        writer: workerIdList1[i][0],
        address: workerIdList1[0][1],
        sheetName:'데이터',
        range:'a2:l2',
      }
      getData1[i] = getSheetData(emptyData);
    }
  }
  
  for (let j=1; j<workerIdList2.length; j++) {
    var copyDataInfo2 = {
      writer: workerIdList2[j][0],
      address: workerIdList2[j][1],
      sheetName:'데이터',
      range:'a2:r2',
    }

    if (copyDataInfo2.address !== '-') {
      getData2[j] = getSheetData(copyDataInfo2);
    } else {
      var emptyData = {
        writer: workerIdList2[j][0],
        address: workerIdList2[0][1],
        sheetName:'데이터',
        range:'a2:l2',
      }
      getData2[j] = getSheetData(emptyData);
    }
    statistics[j-1] = [workerIdList2[j][0], ...upDateTime(), ...getData1[j][0], ...getData2[j][0]]
  }

  var copyDataInfo3 = {
    address:'1EXitttm_MV9xKEBaouVZxeaKPwQAAUn1d9g7YaQ2Lis',
    sheetName:'통계',
    range:'a2:b19',
  }
  
  getData3 = getSheetData(copyDataInfo3);

  for(let k = 0; k < statistics.length; k++) {
    for(let u = 0; u < getData3.length; u++) {
      if(statistics[k][0] === getData3[u][0]) {
        statistics[k].push(getData3[u][1]);
      }
    }
    if(statistics[0].length > statistics[k].length) {
      statistics[k].push(0);
    }
  }

  return statistics;
}

function createSheetAction() {
  var createDataInfo = {
    address:'1HPdmA3P_IWLBYFdhqIHGjAg4FYthoVnCHYqs06EOQlk',
    sheetName:'script(폭력/범죄조장)',
    getColumn: 2
  }
  var workerIdList = getWorkerId(createDataInfo.getColumn);
  for (let i=0; i<workerIdList.length; i++) {
    var createDataInfo = {
      writer: workerIdList[i][0],
      address: workerIdList[i][1],
      sheetName:createDataInfo.sheetName,
    }
    createSheet(createDataInfo);
    Logger.log(createDataInfo.writer);
  }
}

function sheetCopyPaste() {
  // adress에는 복사할 템플릿 시트 id
  // sheetName에는 복사할 템플릿 시트 이름
  // range에는 복사할 영영
  // 여러영역 예시: A10:B10
  // 단일영역 예시: A10
  // getColumn에는 명대사, 페르소나 등 일반 작화 복붙할땐 숫자 '1'
  // getColumn에는 MTS 복붙할땐 숫자 '2'
  var copyDataInfo = {
    address:'1HPdmA3P_IWLBYFdhqIHGjAg4FYthoVnCHYqs06EOQlk',
    sheetName:'데이터',
    range:'a2:C2',
    getColumn: 2
  }

  if(copyDataInfo.sheetName === 'script(성희롱)' ||
    copyDataInfo.sheetName === 'script(인종/지역혐오)' ||
    copyDataInfo.sheetName === 'script(연령혐오)' ||
    copyDataInfo.sheetName === 'script(성혐오)' ||
    copyDataInfo.sheetName === 'script(폭력/범죄조장장)' ||
    copyDataInfo.sheetName === 'script(페르소나)' ||
    copyDataInfo.sheetName === 'script(명대사)' ||
    copyDataInfo.sheetName === 'script(페르소나-코코)' ||
    copyDataInfo.sheetName === 'script(페르소나-마스)') {
    return 0;
  }

  var getData = copySheet(copyDataInfo);
  var workerIdList = getWorkerId(copyDataInfo.getColumn);
  Logger.log(workerIdList)
  for (let i=1; i<workerIdList.length; i++) {
    var pasteDataInfo = {
      writer: workerIdList[i][0],
      address: workerIdList[i][1],
      sheetName:copyDataInfo.sheetName,
      range:copyDataInfo.range,
      data: getData
    }
    pasteSheet(pasteDataInfo);
    Logger.log(pasteDataInfo.writer)
  }
}

/**
 * @param {string} sheetType 복사할 엑셀 형식을 정해줍니다 (현재는 '일반작화', 'MTS작화' 2개 존재)
 */
function getWorkerId(column) {
  var idArray = [];
  var colNumber = column;

  for (let i=0; i < workerInfo.length; i++) {
    idArray[i] = [workerInfo[i][0],workerInfo[i][colNumber]]
  }

  return idArray;
}

/**
 * @param {string} address 구글시트 id값이 들어가야함
 * @param {string} sheetName 구글시트 안의 시트 이름이 들어가야함
 * @param {string} range 셀 영역 설정
 */
function copySheet(datas) {
  var {address,sheetName,range} = datas;
  var ss = SpreadsheetApp.openById(address);
  var sheet = ss.getSheetByName(sheetName);
  var copyData = '';

  if (!/\:/.test(range)) {
    copyData = sheet.getRange(range).getValue();
    var copyFormula = sheet.getRange(range).getFormula();
    if(copyFormula !== "") {
      copyData = copyFormula;
    }
  } else {
    copyData = sheet.getRange(range).getValues();
    var copyFormulas = sheet.getRange(range).getFormulas();

    for(let i=0; i<copyData.length; i++) {
      for(let j=0; j<copyData[i].length; j++) {
        if(copyFormulas[i][j] !== "") {
          copyData[i][j] = copyFormulas[i][j];
        }
      }    
    }
  }
  return copyData;
}

/**
 * @param {string} address 구글시트 id값이 들어가야함
 * @param {string} sheetName 구글시트 안의 시트 이름이 들어가야함
 * @param {string} range 셀 영역 설정
 */
function getSheetData(datas) {
  var {writer,address,sheetName,range} = datas;
  var ss = SpreadsheetApp.openById(address);
  var sheet = ss.getSheetByName(sheetName);
  var copyData = '';

  copyData = sheet.getRange(range).getValues();

  return copyData;
}

/**
 * @param {string} address 구글시트 id값이 들어가야함
 * @param {string} sheetName 구글시트 안의 시트 이름이 들어가야함
 * @param {string} range 셀 영역 설정
 * @param {string} data 붙여넣을 데이터
 */
function pasteSheet(datas) {

  var {address,sheetName,range,data} = datas;
  var ss = SpreadsheetApp.openById(address);
  var sheet = ss.getSheetByName(sheetName);

  if (!/\:/.test(range)) {
    sheet.getRange(range).setValue(data);
  } else {
    sheet.getRange(range).setValues(data);
  }
}

/**
 * @param {string} address 구글시트 id값이 들어가야함
 * @param {string} sheetName 구글시트 안의 시트 이름이 들어가야함
 */
function clearSheet(datas) {
  var {address,sheetName} = datas;
  var ss = SpreadsheetApp.openById(address);
  var sheet = ss.getSheetByName(sheetName);
  sheet.clear();
}

/**
 * @param {string} address 구글시트 id값이 들어가야함
 * @param {string} sheetName 구글시트 안의 시트 이름이 들어가야함
 */
function createSheet(datas) {
  var {address,sheetName} = datas;
  var ss = SpreadsheetApp.openById(address);
  var checkSheet = ss.getSheetByName(sheetName);
  
  if(checkSheet === null) {
    ss.insertSheet(sheetName);
  }
}

/**
 * @param {string} address 구글시트 id값이 들어가야함
 * @param {string} sheetName 구글시트 안의 시트 이름이 들어가야함
 */
// function deleteSheet(datas) {
//   var {address,sheetName} = datas;
//   var ss = SpreadsheetApp.openById(address);
//   var checkSheet = ss.getSheetByName(sheetName);
  
//   if(checkSheet !== null) {
//     ss.deleteSheet(checkSheet);
//     Logger.log('delete success!!');
//   }
// }
