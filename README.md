# node-excel

一个将Excel表格转换成JSON数据的软件,使用nodejs语言开发.  

Tips:
1.可以配置需要导出的数据的Excel名字和对应要导出的Sheet Name.  
2.可以过滤掉空行可以过滤掉以感叹号开头的数据.

## Build

> run build.bat 

## Export

> run export.bat

## Sample

> excel example  

Excel源数据:

![](https://github.com/onelei/node-excel/blob/master/image/excel.png)  

Json数据:

![](https://github.com/onelei/node-excel/blob/master/image/json.png)  

配置文件:

![](https://github.com/onelei/node-excel/blob/master/image/configure.png)  

输入目录:

![](https://github.com/onelei/node-excel/blob/master/image/excelSource.png)  

输出目录:

![](https://github.com/onelei/node-excel/blob/master/image/jsonResult.png)  


> source Path: tables/excel/*

> export Path: tables/json/*

> log Path: tables/output/log.json
