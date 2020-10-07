import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/rtl.xlsx";
  var bytes = File(file).readAsBytesSync();
  //var excel = Excel.createExcel();
  // or
  var excel = Excel.decodeBytes(bytes);
  for (var table in excel.tables.keys) {
    print(table);
    print(excel.tables[table].maxCols);
    print(excel.tables[table].maxRows);
    for (var row in excel.tables[table].rows) {
      print("$row");
    }
  }

  //
  // Change sheet from rtl to ltr and vice-versa
  //
  var sheet1rtl = excel['Sheet1'].isRTL;

  excel['Sheet1'].isRTL = false;

  print('Sheet1: ((previous) isRTL: $sheet1rtl) ---> ((current) isRTL: ${excel['Sheet1'].isRTL})');

  var sheet2rtl = excel['Sheet2'].isRTL;

  excel['Sheet2'].isRTL = true;

  print('Sheet2: ((previous) isRTL: $sheet2rtl) ---> ((current) isRTL: ${excel['Sheet1'].isRTL})');

  CellStyle cellStyle = CellStyle(
    bold: true,
    italic: true,
    fontFamily: getFontFamily(FontFamily.Comic_Sans_MS),
  );

  var sheet = excel['mySheet'];

  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = "Heya How are you I am fine ok goood night";
  cell.cellStyle = cellStyle;

  var cell2 = sheet.cell(CellIndex.indexByString("E5"));
  cell2.value = "Heya How night";
  cell2.cellStyle = cellStyle;

  /// printing cell-type
  print("CellType: " + cell.cellType.toString());

  /// Iterating and changing values to desired type
  for (int row = 0; row < sheet.maxRows; row++) {
    sheet.row(row).forEach((cell) {
      var val = cell.value; //  Value stored in the particular cell

      cell.value = ' My custom Value ';
    });
  }

  excel.rename("mySheet", "myRenamedNewSheet");

  // fromSheet should exist in order to sucessfully copy the contents
  excel.copy('myRenamedNewSheet', 'toSheet');

  excel.rename('oldSheetName', 'newSheetName');

  excel.delete('Sheet1');

  excel.unLink('sheet1');

  sheet = excel['sheet'];

  /// appending rows
  //List<List<String>> list = List.generate(6000, (index) => List.generate(20, (index1) => '$index $index1'));

  //Stopwatch stopwatch = new Stopwatch()..start();
  //list.forEach((row) {
  //  sheet.appendRow(row);
  //});

  //print('doSomething() executed in ${stopwatch.elapsed}');

  sheet.appendRow([8]);
  excel.setDefaultSheet(sheet.sheetName).then((isSet) {
    // isSet is bool which tells that whether the setting of default sheet is successful or not.
    if (isSet) {
      print("${sheet.sheetName} is set to default sheet.");
    } else {
      print("Unable to set ${sheet.sheetName} to default sheet.");
    }
  });

  // Saving the file

  String outputFile = "/Users/kawal/Desktop/rtl.xlsx";
  excel.encode().then((onValue) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
}
