/**
 * Created by onelei on 2017/3/6.
 */

var fs = require('fs');
var node_xj = require("xls-to-json");
var XLSX = require('xlsx');

var excelPath = __dirname+"\\tables\\excel\\";
var outPath = __dirname+"\\tables\\json\\";

/**
 * 遍历给定目录下的所有文件;
*/

function GetAllFiles(path) {
    var fileList = [];
    var dirList = fs.readdirSync(path);
    dirList.forEach(function(item) {
        if (fs.statSync(path + '\\' + item).isDirectory()) {
            GetAllFiles(path + '\\' + item);
        } else {
            //fileList.push(path + '\\' + item);
            fileList.push(item);
        }
    });
    return fileList;
}

var files = GetAllFiles(excelPath);
var errorCount = 0;
console.log("====================== 导表开始 ======================");
for(var i=0;i<files.length;++i){
    // excel name;
    var excelName = files[i];
    var excelPath = excelPath+"\\"+excelName;
    console.log("excel name "+excelName);

    // excel data;
    var excelData = XLSX.readFile(excelPath, null);
    //var excelDataParse = XLSX.parse(excelPath);
    // excel sheet name;
    var sheetNames = excelData['SheetNames'];
    var sheetDatas = excelData['Strings'];
    var sheetDataIndex = 0;


    for(var j=0;j<sheetNames.length;++j){
        var oneSheetName  = sheetNames[j];
        console.log("sheet name "+oneSheetName);

        var oneOutPut = outPath+"\\"+oneSheetName+".json";
        node_xj({
            input: excelPath,  // input xls
            output: oneOutPut, // output json
            sheet: oneSheetName  // specific sheetname
        }, function(err, result) {
            if(err) {
                ++ errorCount;
                console.error(err);
            } else {
                console.log(excelPath +"is over.");
            }
        });
    }

}
console.log("====================== 导表结束 ======================\n"+errorCount+" error.");

function ExcelToJson(excelPath,sheetNames) {
    for(var j=0;j<sheetNames.length;++j){
        var oneSheetName  = sheetNames[j];
        console.log("sheet name "+oneSheetName);

        var oneOutPut = outPath+"\\"+oneSheetName+".json";
        node_xj({
            input: excelPath,  // input xls
            output: oneOutPut, // output json
            sheet: oneSheetName  // specific sheetname
        }, function(err, result) {
            if(err) {
                ++ errorCount;
                console.error(err);
            } else {
                console.log(excelPath +"is over.");
            }
        });
    }
}

/*
var oneExcel = excelPath+"\\test.xlsx";
var oneOutPut = outPath+"\\test.json";
var oneSheetName = "sheet1";

node_xj({
    input: oneExcel,  // input xls
    output: oneOutPut, // output json
    sheet: oneSheetName  // specific sheetname
}, function(err, result) {
    if(err) {
        console.error(err);
    } else {
        console.log(result);
    }
});*/

