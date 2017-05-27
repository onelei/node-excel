/**
 * Created by onelei on 2017/5/25.
 */

var configure = require('./buildconfig.json');
var fs = require('fs');
var XLSX = require('xlsx');

var excelPath = configure.path.source;
var outputPath = configure.path.output;


console.log(" excel path : " +excelPath);

 var oneExcel = excelPath+'\\test.xlsx';
 var oneOutPut = outputPath+'\\test.json';
 var oneSheetName = "sheet1";


var workbook = XLSX.readFile(oneExcel, null);
// 获取 Excel 中所有表名;返回 ['sheet1', 'sheet2']
var sheetNames = workbook.SheetNames;
// 根据表名获取对应某张表
var worksheet = workbook.Sheets[sheetNames[0]];
var headers = {};
var data = [];
var invalidData = [];
var keys = Object.keys(worksheet);

keys.filter(k => k[0]!=='!');// 过滤以 ! 开头的 key

// 遍历所有单元格
keys.forEach(k =>
{
    // 每列;如 A11 中的 A
    var col = k.substring(0, 1);
    // 每行;如 A11 中的 11
    var row = parseInt(k.substring(1));
    // 当前单元格的值
    var value = worksheet[k].v;

    // 保存字段名
    if (row === 3) {
        headers[col] = value;
        return;
    }

    //内容;
    if(row>=4)
    {
        //字符串以!开头的跳过;
        row = row-4;
        // 解析成 JSON

        //if(contains(invalidData,row))
        {
            //return;
        }
        if (!data[row]) {
            //if(value.toString().indexOf('!')==0)
            {
                //invalidData.push(row);
                //console.log(invalidData);
               // return;
            }
            data[row] = {};
        }
        data[row][headers[col]] = value;
    }
});


console.log(data);
fs.writeFileSync(oneOutPut,JSON.stringify(data));
//console.log('导表成功');

function contains(array,key) {
    var i = array.length;
    while (i--) {
        if (array[i] === key) {
            return true;
        }
    }
    return false;
}
