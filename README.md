# excel-by-text
Create complex excel from text file. This program is executable for windows, mac, linux. Language independent.

[日本語の説明はこちら](https://github.com/mikeshimura/excel-by-text/wiki/%E6%97%A5%E6%9C%AC%E8%AA%9E%E8%AA%AC%E6%98%8E)

This program use github.com/tealeg/xlsx for excel generation.

##First part of following Simple report is as follows.  
Columns must be separated by tab(\t).

```go
STA	Sheet1
SN	base	Verdana	10
CS	base	TBLR
SB	TBLR	TBLR	Thin
SF	TBLR	Solid	Blue:20	Yellow
CS	TBLR	TBLR_R
SH	TBLR	Right
SN	TITLE	Arial	24	TBLR	Double
SF	TITLE	Gray125	Black:50	CCCCFF
SI	TITLE	T
SH	TITLE	Center
SN	DATE	Arial	11
SC	DATE	Black:60
CS	TBLR	HEADER
SBL	HEADER	T
SF	HEADER	Solid	Blue:40	Yellow
SB	HEADER	TB	Medium
```

Generated Excel Sample
![Simple1](https://bytebucket.org/mikeshimura/goreport/wiki/image/simple1text.jpg "Simple1")
[text](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/simple1.txt)
[excel](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/simple1.xlsx)

##Execution  
excel-by-text(.exe) -e encoding inputfile

encoding default is UTF8. Accept ShiftJIS, EUCJP.

##Download  
[windows 386](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/excel-by-text_windows_386.exe)  
[windows amd64](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/excel-by-text_windows_amd64.exe)  
[mac 386](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/excel-by-text_darwin_386)  
[mac amd64](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/excel-by-text_darwin_386)  
[linux 386](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/excel-by-text_linux_386)  
[linux amd64](https://bytebucket.org/mikeshimura/goreport/wiki/excel-by-text/excel-by-text_linux_amd64)  

##Commands  
Elements must be separated by tab(\t).

- File Open/Save  
 Create New Excel need no command
 - Existing Excel Open  
  O {filename}  

 - Generated Excel Save  
 W {filename}

- Style  
 - Create New Style     
 SN {stylename} {font} {size} {border} {borderPattern}
