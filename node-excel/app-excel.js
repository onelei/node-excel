/**
 * Created by onelei on 2017/5/31.
 */

var configure = require('../tables/excelconfig.json');
var LOG_PATH = '../tables/output/log.json';
var mLogData = require(LOG_PATH);
var fs = require('fs');
var xlsx = require('node-xlsx');
var path = require('path');

console.log(" 当前目录 : " +__dirname);

var excelPath = configure.path.source;
excelPath = __dirname+'/'+excelPath;
console.log(" excel src path : " +excelPath);

var outputPath = configure.path.output;

outputPath = __dirname+'/'+outputPath;
console.log(" excel out path : " +outputPath);


var log = JSON.stringify(mLogData);
var logJson = JSON.parse(log);
console.log("本地日志信息");
if(logJson.length>0)
{
    console.log(logJson[0]);
}
var exportLog = [];

var excels = configure.excel;
//for(var i=0;i<excels.length;++i)
//用foreach处理异步操作;
var excelIndex = 0;
excels.forEach(function (oneExcel) {
    var data = oneExcel.split(':');
    var excelName = data[0];
    console.log("Excel Name: "+excelName);

    var filePath = excelPath+"\\"+excelName;
    var sheetNames = data[1].split(",");
    fs.stat(filePath,function(err,data){
        //修改时间
        var fileTime = Date.parse(data.mtime);
        console.log(filePath +  "文件的上一次修改时间: "+fileTime);
    //var fileTime = getFileLastModifyTime(filePath,function () {
        excelName = excelName.replace('.xlsx',"");

        var logTime = 0;
        for(var j=0;j<logJson.length;++j)
        {
            if(logJson[j][excelName]!=null)
            {
                logTime = logJson[j][excelName];
                break;
            }
        }
        console.log(excelName+" 文件的log日志时间: "+logTime);

        if(fileTime>logTime)
        {
            console.log("上一次有修改,开始导表");
            praseExcel(filePath,sheetNames);

        }
        else
        {
            console.log("最近没有修改,跳过");
        }

        ++excelIndex;

        var logfile = {};
        logfile[excelName] = fileTime;
        exportLog.push(logfile);
        //console.log(excelIndex);
        //console.log(excels.length);
        if(excelIndex>=excels.length)
        {
            console.log("导出日志");
            var mylogPath = __dirname+'/'+LOG_PATH;
            console.log("log path: "+mylogPath);
            var logData = JSON.stringify(exportLog);
            writeFile(mylogPath,logData);
        }

    });
});

//解析Excel
function praseExcel(filePath,sheetNames)
{
    console.log("文件路径: "+filePath);

    var list = xlsx.parse(filePath);
    if(list==null)
    {
        console.log(filePath+"文件解析失败.");
    }
    console.log("======解析开始======");
    for (var i = 0; i < list.length; i++)
    {
        //每一个sheet下的excel内容;
        var sheetName  = list[i].name;
        //当前的sheet不需要转换,跳过当前的Sheet;
        if(sheetNames.indexOf(sheetName)<0)
        {
            console.log("当前的sheet不需要转换,跳过 Sheet Name: "+sheetName);
            break;
        }
        var excleData = list[i].data;
        if(excleData.length<3)
        {
            console.log("Excel文件不到3行,解析错误. SheetName: "+sheetName);
            return;
        }
        var sheetArray  = [];
        // 第二行是数据类型;
        var typeArray =  excleData[1];
        // 第三行是数据内容;
        var keyArray =  excleData[2];
        for (var j = 3; j < excleData.length ; j++)
        {
            //每一行数据;
            var curData = excleData[j];
            //当前是空行的跳过;
            if(curData.length == 0) continue;
            //以感叹号开头的数据跳过;
            if(curData[0].toString().indexOf('!')==0)
            {
                console.log("跳过当前行, 行数为: "+j+" ID 为 "+curData[0]);
                continue;
            }
            var item = changeObj(curData,typeArray,keyArray);
            sheetArray.push(item);
        }
        //sheet有内容就保存文件;
        if(sheetArray.length >0)
            writeFile(outputPath+"/"+sheetName+".json",JSON.stringify(sheetArray));

        console.log("Sheet Name:  "+sheetName+"  解析结束.");
    }
    console.log(filePath +"  文件解析结束.");

}
//转换数据类型
function changeObj(curData,typeArray,keyArray)
{
    var obj = {};
    for (var i = 0; i < curData.length; i++)
    {
        //字母
        obj[keyArray[i]] = fixValue(curData[i],typeArray[i]);
    }
    return obj;
}

//修复数据类型
function fixValue(value,type)
{
    //空;
    if(value == null || value =="null") return "";
    // int 类型;
    if(type.toString().toLowerCase() =="int".toLowerCase()) return Math.floor(value);
    //默认string类型;
    return value;
}
//写文件
function writeFile(fileName,data)
{
    fs.writeFile(fileName,data,'utf-8',result);
    function result(err)
    {
        if(!err)
        {
            console.log(fileName + " 文件生成成功");
            console.log(Date.now());
        }
        else
        {
            throw err;
            console.log(fileName + " 文件生成失败");
        }
    }
}

function getFileLastModifyTime(fileName,callback)
{
    fs.stat(fileName,function(err,data){
        //修改时间
        var time = Date.parse(data.mtime);
        console.log(fileName +  "文件的上一次修改时间: "+time);
        //执行回调;
        if(callback && typeof(callback) === "function"){
            callback();
        }
    });
}