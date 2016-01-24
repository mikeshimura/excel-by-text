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

-e is optional  
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

- Sheet
 - Add Sheet  
 STA {sheetName}  
 - Set Sheet  
 STS {sheetName}  

 Border pattern and Fill pattern Sample
 ![Sample](https://bytebucket.org/mikeshimura/goreport/wiki/image/parameter2.jpg "Sample")

 [excel](https://bytebucket.org/mikeshimura/goreport/wiki/excel/parameter2.xlsx)  

Color  

- Blue, Yellow etc.  

- FF0000(Red) type RGB expression can be used.  

- Blue:50 Density can be used.

- [defined color text file](https://bytebucket.org/mikeshimura/goreport/wiki/excel/color.txt)
 Color Dictionary (May be registerd)  [www.colordic.org](http://www.colordic.org/)

- Style  
 - Create New Style     
 SN {stylename} {font} {size} {border} {borderPattern}  
 // border "T" top  "B" bottom  "L" left  "R" right  
 //border and borderPattern are optional  

 - Copy Style  
 CS {FromStyleName} {ToStyleName}  

 - Set Font Name  
 SFN {stylename} {fontName}  

 - Set Font Size    
 SFS {stylename} {fontSize}  

 - Set Font Color    
 SFS {stylename} {color}  

 - Set Italic    
 SI {stylename} {bool}  
 //bool T (True) or F (False)  

 - Set Bold     
 SBL {stylename} {bool}  

 - Set Underline   
 SU {stylename} {bool}  

 - Set Border    
 SB {stylename} {border} {borderPattern}  
 // border "T" top  "B" bottom  "L" left  "R" right  

 - Set Fill   
 SF {stylename} {pattern} {fgColor} {bgColor}  

 - Set Horizontal Align   
 SH {alignment}  
 //Left, Center, Right, Justify, Distributed, CenterContinuous, Fill, General  

 - Set Vertical Align   
 SV {alignment}  
 //Top, Center, Bottom, Justify, Distributed  

- Cell Value and Format  

 - Set Column Width  
 CW {startCol} {endCol} {width}  

 - Merge  
 M {rowno} {colno} {toRowno} {toColno}  

 - Set Format  
 FS {rowno} {colno} {format}  

 - Set String  
 S {rowno} {colno} {content}  

 - Set Number
 N {rowno} {colno} {value}  

 - Set Number Format  
 NF {rowno} {colno} {value} {format}  

 - Set Date
 D {rowno} {colno} {value}  
 // date value format yyyy/mm/dd

 - Set Date Format  
 DF {rowno} {colno} {value} {format}

 - Set DateTime
 DT {rowno} {colno} {value}  
 // datetime value format yyyy/mm/dd hh:mm:ss  

 - Set DateTime Format  
 DTF {rowno} {colno} {value} {format}

 - Set Formula
 F {rowno} {colno} {formula}  

 - Set formula Format  
 FF {rowno} {colno} {formula}  {format}
