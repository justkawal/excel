## Donate (Be the First to Donate)

Please consider donating if you think excel is helpful to you or that my work is valuable. I will be happy if you can help me upgrade my lazy laptop. [Paypal Me on paypal.me/kawal7415](https://www.paypal.me/kawal7415)

[![GitHub issues](https://img.shields.io/github/issues/kawal7415/excel)](https://github.com/kawal7415/excel/issues)[![GitHub forks](https://img.shields.io/github/forks/kawal7415/excel)](https://github.com/kawal7415/excel/network)[![GitHub stars](https://img.shields.io/github/stars/kawal7415/excel)](https://github.com/kawal7415/excel/stargazers)

# Excel

[Excel](https://www.pub.dev/packages/excel) is a flutter and dart library for creating and updating excel-sheets for XLSX files.

## Usage

### Adding dependency in pubspec.yaml

````dart
    
dependencies:
    excel: ^1.0.3
        
````

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
    var updater = Excel.createExcel(); //automatically creates 3 empty sheets Sheet1, Sheet2 and Sheet3 
     
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
      * updater.updateCell('sheetName', cell, value, options?);
      * if sheet === 'sheetName' does not exist in updater, it will be created automatically after calling updateCell method
      * cell can be identified with Cell Address or by 2D array having row and column Index;
      * Cell options are optional
      */
      
      var sheet = 'SheetName';
      
      //update cell with cellAddress
      updater.updateCell(sheet, CellIndex.indexByString("A1"), "Here value of A1");
        
      //update cell with row and column index
      updater.updateCell(sheet, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0), "Here value of C1");
        
      //update cell and it's background color
      deocder.updateCell(sheet, CellIndex.indexByString("A2"), "Here value of A2", backgroundColorHex: "#1AFF1A")
      
      //update alignment
      updater.updateCell(sheet, CellIndex.indexByString("E5"), "Here value of E5", horizontalAlign: HorizontalAlign.Right);
    
   ````
### Cell Options
key | description
------------ | -------------
 fontColorHex | Font Color eg. "#0000FF"
 backgroundColorHex | Background color of cell eg. "#faf487"
 wrap | Text wrapping ````enum TextWrapping { WrapText, Clip }```` eg. TextWrapping.Clip
 verticalAlign | align text vertically ````enum VerticalAlign { Top, Middle, Bottom }```` eg. VerticalAlign.Top
 horizontalAlign | align text horizontally ````enum HorizontalAlign { Left, Center, Right }```` eg. HorizontalAlign.Right


 ### Merge Cells
 
 ````dart
     /* 
     * updater.merge('sheetName', starting_cell, ending_cell, 'customValue');
     * sheet === 'sheetName' in which merging of rows and columns is to be done
     * starting_cell and ending_cell can be identified with Cell Address or by 2D array having row and column Index;
     * customValue is optional
     */
 
      updater.merge(sheet, CellIndex.indexByString("A1"), CellIndex.indexByString("E4"), customValue: "Put this text after merge");
    
   ````
   
 ### Get Merged Cells List
 
 ````dart
      // Check which cells are merged
 
      updater.getSpannedItems(sheet).forEach((cells) {
        print("Merged:" + cells.toString());
      });
    
   ````
   
 ### Un-Merge Cells
 
 ````dart
     /* 
     * updater.unMerge(sheet, cell);
     * sheet === 'sheetName' in which un-merging of rows and columns is to be done
     * cell should be identified with string only with an example as "A1:E4"
     * to check if "A1:E4" is un-merged or not
     * call the method updater.getSpannedItems(sheet); and verify that it is not present in it.
     */
 
      updater.unMerge(sheet, "A1:E4");
    
   ````
 
  ### Get Default Opening Sheet
 
 ````dart
     /* 
     * Asynchronous method which returns the name of the default sheet
     * updater.getDefaultSheet();
     */
 
      updater.getDefaultSheet().then((value) {
        print("Default Sheet:" + value.toString());
      });
      
      or
      
      var defaultSheet = await updater.getDefaultSheet();
      print("Default Sheet:" + defaultSheet.toString());
    
   ````
   
  ### Set Default Opening Sheet
 
 ````dart
     /* 
     * Asynchronous method which sets the name of the default sheet
     * returns bool if successful then true else false
     * updater.setDefaultSheet(sheet);
     * sheet = 'SheetName'
     */
 
      updater.setDefaultSheet(sheet).then((isSet) {
        if (isSet) {
            print("$sheet is set to default sheet.");
        } else {
            print("Unable to set $sheet to default sheet.");
        }
      });
      
      or
      
      var isSet = await updater.setDefaultSheet(sheet);
      if (isSet) {
        print("$sheet is set to default sheet.");
      } else {
        print("Unable to set $sheet to default sheet.");
      }
    
   ````
   
 ### Saving XLSX File
 
 ````dart
      // Save the Changes in file

      updater.encode().then((onValue) {
        File(join("Path_to_destination/excel.xlsx"))
        ..createSync(recursive: true)
        ..writeAsBytesSync(onValue);
    });
    
   ````

## Features coming in next version
On-going implementation for future:
- Formulas
- Font Family
- Text Size
- Italic
- Underline
- Bold

## Important:
For XLSX format, this implementation only supports native Excel format for date, time and boolean type conversion.
In other words, custom format for date, time, boolean aren't supported and also the files exported from LibreOffice as well.

## Paypal Me

[Paypal Me](https://www.paypal.me/kawal7415)
