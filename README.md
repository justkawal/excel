# Excel
  
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
   <a href="https://github.com/justkawal/excel/issues">  
    <img src="https://img.shields.io/github/issues/justkawal/excel"  
      alt="Issue" />  
  </a> 
   <a href="https://github.com/justkawal/excel/network">  
    <img src="https://img.shields.io/github/forks/justkawal/excel"  
      alt="Forks" />  
  </a> 
   <a href="https://github.com/justkawal/excel/stargazers">  
    <img src="https://img.shields.io/github/stars/justkawal/excel"  
      alt="Stars" />  
  </a>
  <br>
  <br>
 
 [Excel](https://www.pub.dev/packages/excel) is a flutter and dart library for reading, creating and updating excel-sheets for XLSX files.



# Table of Contents
  - [Installing](#lets-get-started)
  - [Usage](#usage)
    * [Breaking Changes](#breaking-changes-for-those-moving-from-108-and-below-------109-and-above-versions)
    * [Imports](#imports)
    * [Read xlsx file](#read-xlsx-file)
    * [Read xlsx file from Asset Folder](#read-xlsx-from-flutters-asset-folder)
    * [Create xlsx file](#create-new-xlsx-file)
    * [Update Cell](#update-cell-values)
      + [Cell Style options](#cell-style-options)
    * [Copy Sheet](#copy-sheet-contents-to-another-sheet)
    * [Rename Sheet](#rename-sheet)
    * [Delete Sheet](#delete-sheet)
    * [Link Sheet](#link-sheet)
    * [Un-Link Sheet](#un-link-sheet)
    * [Merge Cells](#merge-cells)
    * [Get Merged Cells](#get-merged-cells-list)
    * [Un-Merge Cells](#un-merge-cells)
    * [Find and Replace](#find-and-replace)
    * [Insert Row Iterables](#insert-row-iterables)
      + [Iterables Options](#iterable-options)
    * [Append Row](#append-row)
    * [Get default sheet](#get-default-opening-sheet)
    * [Set default sheet](#set-default-opening-sheet)
    * [Saving](#saving-xlsx-file)
  - [Frequent Issues](#frequent-issues)
  - [Upcoming Features](#features-coming-in-next-version)
  - [Donate (Be the First one)](#donate-be-the-first-one)

# Lets Get Started

### 1. Depend on it
Add this to your package's `pubspec.yaml` file:

```yaml
dependencies:
  excel: ^1.1.5
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

### Breaking Changes for those moving from 1.0.8 and below ---->  1.0.9 and above versions

The necessary changes to be made to updateCell function in order to prevent the code from breaking.

````dart
    
    excel.updateCell('SheetName', CellIndex.indexByString("A2"), "Here value", backgroundColorHex: "#1AFF1A", horizontalAlign: HorizontalAlign.Right);
    
    // Now in the above code wrap the optional arguments with CellStyle() and pass it to optional cellStyle parameter.
    // So the resulting code will look like
    
    excel.updateCell('SheetName', CellIndex.indexByString("A2"), "Here value", cellStyle: CellStyle( backgroundColorHex: "#1AFF1A", horizontalAlign: HorizontalAlign.Right ) );
    

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
    var excel = Excel.decodeBytes(bytes);
    
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

### Create New XLSX File
    
````dart
    var excel = Excel.createExcel(); // automatically creates 1 empty sheet: Sheet1
    
````

 ### Update Cell values
 
 ````dart
     /* 
      * sheetObject.updateCell(cell, value, { CellStyle (Optional)});
      * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
      * cell can be identified with Cell Address or by 2D array having row and column Index;
      * Cell Style options are optional
      */
      
      Sheet sheetObject = excel['SheetName'];
      
      CellStyle cellStyle = CellStyle(backgroundColorHex: "#1AFF1A", fontFamily : getFontFamily(FontFamily.Calibri));
      
      cellStyle.underline = Underline.Single; // or Underline.Double
      
      
      var cell = sheetObject.cell(CellIndex.indexByString("A1"));
      cell.value = 8; // dynamic values support provided;
      cell.cellStyle = cellStyle;
      
      // printing cell-type
      print("CellType: "+ cell.cellType.toString());
      
      ///
      /// Inserting and removing column and rows
      
      // insert column at index = 8
      sheetObject.insertColumn(8);
      
      // remove column at index = 18
      sheetObject.removeColumn(18);
      
      // insert row at index = 82
      sheetObject.removeRow(82);
      
      // remove row at index = 80
      sheetObject.removeRow(80);
      
      
````

### Cell-Style Options
key | description
------------ | -------------
 fontFamily | eg. getFontFamily(``FontFamily.Arial``) or getFontFamily(``FontFamily.Comic_Sans_MS``) ``There is total 182 Font Families available for now``
 fontSize | specify the font-size as integer eg. fontSize = 15
 bold | makes text bold - when set to ``true``, by-default it is set to ``false``
 italic | makes text italic - when set to ``true``, by-default it is set to ``false``
 underline | Gives underline to text ``enum Underline { None, Single, Double }`` eg. Underline.Single, by-default it is set to Underline.None
 fontColorHex | Font Color eg. "#0000FF"
 rotation | rotation of text eg. 50, rotation varies from ``-90 to 90``, with including ``90`` and ``-90``
 backgroundColorHex | Background color of cell eg. "#faf487"
 wrap | Text wrapping ``enum TextWrapping { WrapText, Clip }`` eg. TextWrapping.Clip
 verticalAlign | align text vertically ``enum VerticalAlign { Top, Center, Bottom }`` eg. VerticalAlign.Top
 horizontalAlign | align text horizontally ``enum HorizontalAlign { Left, Center, Right }`` eg. HorizontalAlign.Right


### Make sheet RTL
 
````dart
 
     /*
      * set rtl to true for making sheet to right-to-left
      * default value of rtl = false ( which means the fresh or default sheet is ltr )
      * 
      */
      
      var sheetObject = excel['SheetName'];
      sheetObject.rtl = true;
      
````
   
### Copy sheet contents to another sheet
 
 ````dart
 
     /*
      * excel.copy(String 'existingSheetName', String 'anotherSheetName');
      * existingSheetName should exist in excel.tables.keys in order to successfully copy
      * if anotherSheetName does not exist then it will be automatically created.
      *
      */
      
      excel.copy('existingSheetName', 'anotherSheetName');
      
````

### Rename sheet
 
 ````dart
 
     /*
      * excel.rename(String 'existingSheetName', String 'newSheetName');
      * existingSheetName should exist in excel.tables.keys in order to successfully rename
      *
      */
      
      excel.rename('existingSheetName', 'newSheetName');
      
````
   
### Delete sheet
 
 ````dart
 
     /*
      * excel.delete(String 'existingSheetName');
      * (existingSheetName should exist in excel.tables.keys) and (excel.tables.keys.length >= 2), in order to successfully delete.
      * 
      */
      
      excel.delete('existingSheetName');
      
````
   
### Link sheet
 
 ````dart
 
     /*
      * excel.link(String 'sheetName', Sheet sheetObject);
      * 
      * Any operations performed on (object of 'sheetName') or sheetObject then the operation is performed on both.
      * if 'sheetName' does not exist then it will be automatically created and linked with the sheetObject's operation.
      *
      */
      
      excel.link('sheetName', sheetObject);
      
   ````
   
### Un-Link sheet
 
 ````dart
 
     /*
      * excel.unLink(String 'sheetName');
      * In order to successfully unLink the 'sheetName' then it must exist in excel.tables.keys
      *
      */
      
      excel.unLink('sheetName');
      
      // After calling the above function be sure to re-make a new reference of this.
      
      Sheet unlinked_sheetObject = excel['sheetName'];
      
````

 ### Merge Cells
 
 ````dart
    /* 
     * sheetObject.merge(CellIndex starting_cell, CellIndex ending_cell, dynamic 'customValue');
     * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
     * starting_cell and ending_cell can be identified with Cell Address or by 2D array having row and column Index;
     * customValue is optional
     */
 
      sheetObject.merge(CellIndex.indexByString("A1"), CellIndex.indexByString("E4"), customValue: "Put this text after merge");
     
````
   
 ### Get Merged Cells List
 
 ````dart
      // Check which cells are merged
      
      sheetObject.spannedItems.forEach((cells) {
        print("Merged:" + cells.toString());
      });
    
````
   
 ### Un-Merge Cells
 
 ````dart
    /* 
     * sheetObject.unMerge(cell);
     * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
     * cell should be identified with string only with an example as "A1:E4".
     * to check if "A1:E4" is un-merged or not
     * call the method excel.getMergedCells(sheet); and verify that it is not present in it.
     */
      
      sheetObject.unMerge("A1:E4");
    
````
   
 ### Find and Replace
 
 ````dart
    /* 
     * int replacedCount = sheetObject.findAndReplace(source, target);
     * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
     * source is the string or ( User's Custom Pattern Matching RegExp )
     * target is the string which is put in cells in place of source
     * 
     * it returns the number of replacements made
     */
      
      int replacedCount = sheetObject.findAndReplace(Flutter', 'Google');
      
````
   
 ### Insert Row Iterables
 
 ````dart
     /* 
      * sheetObject.insertRowIterables(list-iterables, rowIndex, iterable-options?);
      * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
      * list-iterables === list of iterables which has to be put in specific row
      * rowIndex === the row in which the iterables has to be put
      * Iterable options are optional
      */
      
      /// It will put the list-iterables in the 8th index row
      List<String> dataList = ["Google", "loves", "Flutter", "and", "Flutter", "loves", "Google"];
      
      sheetObject.insertRowIterables(dataList, 8);
    
````

### Iterable Options
key | description
------------ | -------------
 startingColumn | starting column index from which list-iterables should be started
 overwriteMergedCells | overwriteMergedCells is by-defalut set to ```true```, when set to ```false``` it will stop over-write and will write only in unique cells
   
 ### Append Row
 
 ````dart
    /* 
     * sheetObject.appendRow(list-iterables);
     * sheetObject created by calling - // Sheet sheetObject = excel['SheetName'];
     * list-iterables === list of iterables
     */
     
      sheetObject.appendRow(["Flutter", "till", "Eternity"]);
    
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
 
# Frequent Issues

### Having Trouble using excel i.e: ````every version of some_other_library depends on version..... blah blah blah````

```Similar issues have been solved in the past :``` [#26](https://github.com/justkawal/excel/issues/26), [#25](https://github.com/justkawal/excel/issues/25), [#11](https://github.com/justkawal/excel/issues/11)


## Features coming in next version
On-going implementation for future:
- Formulas
- Conversion to PDF

#### Also checkout our other libraries: 
  - Protect **···············**> [Protect](https://www.github.com/justkawal/protect)
  - Text Animations **··**> [AnimatedText](https://www.github.com/justkawal/animated_text)
  - Translations **·······**> [Arb Translator](https://www.github.com/justkawal/arb_translator)

### Donate
Ooooops, My laptop is **slow**, but I'm not.
  - [Paypal](https://www.paypal.me/kawal7415)
  - Not having Paypal account ?? [Join Now](https://www.paypal.com/in/flref?refBy=Pzpaa7qp041602067472432) and both of us could earn **`$10`**
