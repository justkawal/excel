## Donate (Be the First to Donate)

Please consider donating if you think excel is helpful to you or that my work is valuable. I will be happy if you can help me buy a [cup of coffee. Paypal Me ](https://www.paypal.me/kawal7415) on paypal.me/kawal7415

[![GitHub issues](https://img.shields.io/github/issues/kawal7415/excel)](https://github.com/kawal7415/excel/issues)[![GitHub forks](https://img.shields.io/github/forks/kawal7415/excel)](https://github.com/kawal7415/excel/network)[![GitHub stars](https://img.shields.io/github/stars/kawal7415/excel)](https://github.com/kawal7415/excel/stargazers)

# Excel

Excel is a flutter and dart library for creating and updating excel-sheets for XLSX files.

## Usage

### Imports

````dart
    import 'dart:io';
    import 'package:path/path.dart';
    import 'package:excel/excel.dart';
    
````
### Read XLSX File

````dart
    var file = "Path_to_pre_existing_Excel_File/excel_file.xlsx";
    var bytes = File(file).readAsBytesSync();
    var updater = Excel.decodeBytes(bytes, update: true);
    
    for (var table in updater.tables.keys) {
      print(table); //sheet Name
      print(updater.tables[table].maxCols);
      print(updater.tables[table].maxRows);
      for (var row in updater.tables[table].rows) {
        print("$row");
      }
    }
    
````
### Create XLSX File
    
````dart
    var updater = Excel.createExcel(); //automatically creates 3 sheets Sheet1, Sheet2 and Sheet3 
     
    //find desired sheet name in updater/file;
    for (var tableName in updater.tables.keys) {
      if( desiredSheetName.toString() == tableName.toString() ){
        sheet = tableName.toString();
        break;
       }
    }
    
````
 ### Update Cell values
 
 ````dart
      /* 
      * updater.updateCell('sheetName',cell,value,options?);
      * if sheet === 'sheetName' does not exist in updater, it will create automatically after calling updateCell method
      * cell can be identified  with Cell Address or by 2D array having row and column Index;
      * Cell options are optional
      */
      
      
      //update cell with cellAddress
      updater.updateCell(sheet, CellIndex.indexByString("A1"), "Here value of A1");
        
      //update cell with row and column index
      updater.updateCell(sheet, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),"Here value of C1");
        
      //update cell and it's background color
      deocder.updateCell(sheet, CellIndex.indexByString("A2"), "Here value of A2",backgroundColorHex: "#1AFF1A")
      
      //update alignment
      updater.updateCell(sheet, CellIndex.indexByString("E5"), "Here value of E5",horizontalAlign: HorizontalAlign.Right);
    
      // Save the Changes in file

      updater.encode().then((onValue) {
        File(join("Path_to_destination/excel.xlsx"))
        ..createSync(recursive: true)
        ..writeAsBytesSync(onValue);
    });
    
   ````
### Cell Options
key | description
------------ | -------------
 fontColorHex | Font Color eg. "#0000FF"
 backgroundColorHex | Background color of cell eg. "#faf487"
 wrap | Text wrapping ````enum TextWrapping { WrapText, Clip }```` eg. TextWrapping.Clip
 verticalAlign | align text vertically ````enum VerticalAlign { Top, Middle, Bottom }```` eg. VerticalAlign.Top
 horizontalAlign | align text horizontally ````enum HorizontalAlign { Left, Center, Right }```` eg. HorizontalAlign.Right

    

## Features coming in next version
On-going implementation for future:
- Formulas
- Font Family
- Text Size

## Important:
For XLSX format, this implementation only supports native Excel format for date, time and boolean type conversion.
In other words, custom format for date, time, boolean aren't supported and also the files exported from LibreOffice as well.

## Paypal Me

[Paypal Me](https://www.paypal.me/kawal7415)
