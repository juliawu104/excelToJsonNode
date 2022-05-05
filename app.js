const xlsx = require("xlsx");
const fs = require("fs");
const excelConfig = {
    entry:'./google_sheet.xlsx',  //讀取excel位置
    sheetNo:0, //讀取excel的哪張sheet
    outputDir:'./',  //導出位置
    outputName:'google_sheet', //導出檔案名稱
}

// 讀檔案
const wb = xlsx.readFile(excelConfig.entry);

// 讀檔案中的sheet轉檔案出來
let data = parseToJson(wb);
function parseToJson(wb) {
  let data = {};
  // 獲取第一個sheet
  let sheetName = wb.SheetNames[excelConfig.sheetNo];
  let sheet = wb.Sheets[sheetName];
  // 獲取第一個sheet

  // 獲取sheet資料範圍
  let range = xlsx.utils.decode_range(sheet["!ref"]);
  let { c: startC, r: startR } = range.s;
  let { c: endC, r: endR } = range.e;
  let lengthC = endC - startC;
  let lengthR = endR - startR;
  // 獲取sheet資料範圍

  // 轉成自己要的格式
  data = getMultiValueArray(
    sheetName,
    sheet,
    range,
    startC,
    startR,
    endC,
    endR,
    lengthC,
    lengthR
  );
  // 轉成自己要的格式
  return data;
}

// 核心邏輯
/**
 * 在這邊客製化你的excel轉出來的檔案即可
 * @param {*} sheetName sheetName
 * @param {*} sheet sheet
 * @param {*} range 獲取有值X,Y的範圍
 * @param {*} startC excel有值的最左邊
 * @param {*} startR excel有值的最上面
 * @param {*} endC excel有值的最右邊
 * @param {*} endR excel有值的最下面
 * @param {*} lengthC excel有值的左到右長度
 * @param {*} lengthR excel有值的上到下長度
 * @returns 
 */
function getMultiValueArray(
  sheetName,
  sheet,
  range,
  startC,
  startR,
  endC,
  endR,
  lengthC,
  lengthR
) {
  let result = [];
  // Handle Basic data
  // Basic.xlsx
  // for (let r = startR; r <= endR; r++) {
  //   let obj = {};
  //   for(let c = startC; c <= endC; c++){
  //     let key = getCellValue(sheet, range, c, 0);
  //     obj[key] = getCellValue(sheet, range, c, r)
  //   }
  //   result.push(obj);
  // }
  // result.shift()

  // Handle Date data formate
  // google_sheet.xlsx
  const dateMode = wb.Workbook.WBProps.date1904;
  for (let r = startR; r <= endR; r++) {
    let obj = {};
    for(let c = startC; c <= endC; c++){
      let key = getCellValue(sheet, range, c, 0);
      obj[key] = getCellValue(sheet, range, c, r)
      if(c == 2){
        let val = getCellValue(sheet, range, c, r);
        obj[key] = xlsx.SSF.format('YYYY/MM/DD', val, { date1904: dateMode })
      }
    }
    result.push(obj);
  }
  result.shift()

 
  
  // console.log(xlsx.SSF.format('YYYY-MM-DD', val, { date1904: dateMode }))
  // 2022-02-02



  return { data: result };
}

// 獲取單元格資料
/**
 * 獲取excel指定單元格位置的資料
 * @param {*} sheet 哪張sheet
 * @param {*} range 可以獲取有值的範圍
 * @param {*} x 從0開始
 * @param {*} y 從0開始
 * @returns 
 */
function getCellValue(sheet, range, x, y) {
  const position = xlsx.utils.encode_cell({
    c: range.s.c + x,
    r: range.s.r + y,
  });
  debugger;
  return sheet[position] ? sheet[position].v : "";
}

// 導出位置
fs.writeFileSync(excelConfig.outputDir + '/' + excelConfig.outputName + '.json', JSON.stringify(data, null, 2));
