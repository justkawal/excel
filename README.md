## Donate (Be the first one)

I can code faster but my laptop is **lazy**. Help me upgrade it. Please consider donating if you think excel is helpful. 

[Paypal Me on paypal.me/kawal7415](https://www.paypal.me/kawal7415)

  <a href="https://flutter.io">  
    <img src="https://img.shields.io/badge/Platform-Flutter-yellow.svg"  
      alt="Platform" />  
  </a> 
   <a href="https://pub.dartlang.org/packages/excel">  
    <img src="https://img.shields.io/pub/v/excel.svg"  
      alt="Pub Package" />  
  </a>
   <a href="https://opensource.org/licenses/MIT">  
    <img src="https://img.shields.io/badge/License-MIT-red.svg"  
      alt="License: MIT" />  
  </a>  
   <a href="https://www.paypal.me/kawal7415">  
    <img src="https://img.shields.io/badge/Donate-PayPal-green.svg"  
      alt="Donate" />  
  </a>
   <a href="https://github.com/kawal7415/excel/issues">  
    <img src="https://img.shields.io/github/issues/kawal7415/excel"  
      alt="Issue" />  
  </a> 
   <a href="https://github.com/kawal7415/excel/network">  
    <img src="https://img.shields.io/github/forks/kawal7415/excel"  
      alt="Forks" />  
  </a> 
   <a href="https://github.com/kawal7415/excel/stargazers">  
    <img src="https://img.shields.io/github/stars/kawal7415/excel"  
      alt="Stars" />  
  </a>

# Excel

[Excel](https://www.pub.dev/packages/excel) is a flutter and dart library for creating and updating excel-sheets for XLSX files.


# Installing

### 1. Depend on it
Add this to your package's `pubspec.yaml` file:

```yaml
dependencies:
  excel: ^1.0.7
```

### 2. Install it

You can install packages from the command line:

with `pub`:

```css
$  pub get
```

with `Flutter`:

```css
$  flutter packages get
```

### 3. Import it

Now in your `Dart` code, you can use: 

````dart
    import 'package:excel/excel.dart';
````


# Usage

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
    var excel = Excel.decodeBytes(bytes, update: true);
    
    for (var table in excel.tables.keys) {
      print(table); //sheet Name
      print(excel.tables[table].maxCols);
      print(excel.tables[table].maxRows);
      for (var row in excel.tables[table].rows) {
        print("$row");
      }
    }
    
````

### Read XLSX from Flutter's Asset Folder

````dart
    import 'package:flutter/services.dart' show ByteData, rootBundle;
    
    /* Your blah blah code here */
    
    ByteData data = await rootBundle.load("assets/existing_excel_file.xlsx");
    var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
    var excel = Excel.decodeBytes(bytes, update: true);
        
    for (var table in excel.tables.keys) {
      print(table); //sheet Name
      print(excel.tables[table].maxCols);
      print(excel.tables[table].maxRows);
      for (var row in excel.tables[table].rows) {
        print("$row");
      }
    }
    
````

### Create XLSX File
    
````dart
    var excel = Excel.createExcel(); //automatically creates 3 empty sheets Sheet1, Sheet2 and Sheet3 
     
    //find desired sheet name in excel/file;
    for (var tableName in excel.tables.keys) {
      if( desiredSheetName.toString() == tableName.toString() ){
        sheet = tableName.toString();
        break;
       }
    }
    
````
 ### Update Cell values
 
 ````dart
      /* 
      * excel.updateCell('sheetName', cell, value, options?);
      * if sheet === 'sheetName' does not exist in excel, it will be created automatically after calling updateCell method
      * cell can be identified with Cell Address or by 2D array having row and column Index;
      * Cell options are optional
      */
      
      var sheet = 'SheetName';
      
      //update cell with cellAddress
      excel.updateCell(sheet, CellIndex.indexByString("A1"), "Here value of A1");
        
      //update cell with row and column index
      excel.updateCell(sheet, CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0), "Here value of C1");
        
      //update cell and it's background color
      excel.updateCell(sheet, CellIndex.indexByString("A2"), "Here value of A2", backgroundColorHex: "#1AFF1A")
      
      //update alignment
      excel.updateCell(sheet, CellIndex.indexByString("E5"), "Here value of E5", horizontalAlign: HorizontalAlign.Right);

      // Insert column at index = 17;
      excel.insertColumn(sheet, 17);
    
      // Remove column at index = 2
      excel.removeColumn(sheet, 2);
    
      // Insert row at index = 2;
      excel.insertRow(sheet, 2);
    
      // Remove row at index = 17
      excel.removeRow(sheet, 2);
    
   ````
### Cell Options
key | description
------------ | -------------
 fontColorHex | Font Color eg. "#0000FF"
 backgroundColorHex | Background color of cell eg. "#faf487"
 wrap | Text wrapping ````enum TextWrapping { WrapText, Clip }```` eg. TextWrapping.Clip
 verticalAlign | align text vertically ````enum VerticalAlign { Top, Center, Bottom }```` eg. VerticalAlign.Top
 horizontalAlign | align text horizontally ````enum HorizontalAlign { Left, Center, Right }```` eg. HorizontalAlign.Right


 ### Merge Cells
 
 ````dart
     /* 
     * excel.merge('sheetName', starting_cell, ending_cell, 'customValue');
     * sheet === 'sheetName' in which merging of rows and columns is to be done
     * starting_cell and ending_cell can be identified with Cell Address or by 2D array having row and column Index;
     * customValue is optional
     */
 
      excel.merge(sheet, CellIndex.indexByString("A1"), CellIndex.indexByString("E4"), customValue: "Put this text after merge");
    
   ````
   
 ### Get Merged Cells List
 
 ````dart
      // Check which cells are merged
 
      excel.getMergedCells(sheet).forEach((cells) {
        print("Merged:" + cells.toString());
      });
    
   ````
   
 ### Un-Merge Cells
 
 ````dart
     /* 
     * excel.unMerge(sheet, cell);
     * sheet === 'sheetName' in which un-merging of rows and columns is to be done
     * cell should be identified with string only with an example as "A1:E4"
     * to check if "A1:E4" is un-merged or not
     * call the method excel.getMergedCells(sheet); and verify that it is not present in it.
     */
 
      excel.unMerge(sheet, "A1:E4");
    
   ````
   
 ### Find and Replace
 
 ````dart
     /* 
     * int replacedCount = excel.findAndReplace(sheetName, source, target);
     * sheet === 'sheetName' in which replacement is to be done
     * source is the string or ( User's Custom Pattern Matching RegExp )
     * target is the string which is put in cells in place of source
     * 
     * it returns the number of replacements made
     */
 
      int replacedCount = excel.findAndReplace(sheet, 'Flutter', 'Google');
      
      or
      
      int replacedCount = excel.findAndReplace(sheet, RegExp('your blah blah important regexp pattern'), 'Google');
      print("Replaced Count:" + replacedCount.toString());
    
   ````
   
 ### Insert Row Iterables
 
 ````dart
      /* 
      * excel.insertRowIterables(sheet, list-iterables, rowIndex, iterable-options?);
      * sheet === 'sheetName'
      * list-iterables === list of iterables which has to be put in specific row
      * rowIndex === the row in which the iterables has to be put
      * Iterable options are optional
      */
      
      /// It will put the list-iterables in the 8th index row
      List<String> dataList = ["Google", "loves", "Flutter", "and", "Flutter", "loves", "Google"];
      excel.insertRowIterables(sheet, dataList, 8);
    
   ````

### Iterable Options
key | description
------------ | -------------
 startingColumn | starting column index from which list-iterables should be started
 overwriteMergedCells | overwriteMergedCells is by-defalut set to ```true```, when set to ```false``` it will stop over-write and will write only in unique cells
   
 ### Append Row
 
 ````dart
     /* 
     * excel.appendRow(sheetName, list-iterables);
     * sheet === 'sheetName' in which the list-iterables is to be put in the last available row.
     * list-iterables === list of iterables
     * 
     */
     
      excel.appendRow(sheet, ["Flutter", "till", "Eternity"]);
    
   ````
 
### Get Default Opening Sheet
 
 ````dart
     /* 
     * Asynchronous method which returns the name of the default sheet
     * excel.getDefaultSheet();
     */
 
      excel.getDefaultSheet().then((value) {
        print("Default Sheet:" + value.toString());
      });
      
      or
      
      var defaultSheet = await excel.getDefaultSheet();
      print("Default Sheet:" + defaultSheet.toString());
    
   ````
   
### Set Default Opening Sheet
 
 ````dart
     /* 
     * Asynchronous method which sets the name of the default sheet
     * returns bool if successful then true else false
     * excel.setDefaultSheet(sheet);
     * sheet = 'SheetName'
     */
 
      excel.setDefaultSheet(sheet).then((isSet) {
        if (isSet) {
            print("$sheet is set to default sheet.");
        } else {
            print("Unable to set $sheet to default sheet.");
        }
      });
      
      or
      
      var isSet = await excel.setDefaultSheet(sheet);
      if (isSet) {
        print("$sheet is set to default sheet.");
      } else {
        print("Unable to set $sheet to default sheet.");
      }
    
   ````
   
 ### Saving XLSX File
 
 ````dart
      // Save the Changes in file

      excel.encode().then((onValue) {
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

## Help us to keep going.

[![Donate with PayPal](https://github.com/kawal7415/excel/blob/master/paypal_png.png)](https://www.paypal.me/kawal7415)
