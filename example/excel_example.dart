import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  //var file = "/Users/kawal/Desktop/excel/test/test_resources/example.xlsx";
  //var bytes = File(file).readAsBytesSync();
  var excel = Excel.createExcel();
  // or
  //var excel = Excel.decodeBytes(bytes);

  ///
  ///
  /// reading excel file values
  ///
  ///
  for (var table in excel.tables.keys) {
    print(table);
    print(excel.tables[table]!.maxCols);
    print(excel.tables[table]!.maxRows);
    for (var row in excel.tables[table]!.rows) {
      print("${row.map((e) => e?.value)}");
    }
  }

  ///
  /// Change sheet from rtl to ltr and vice-versa i.e. (right-to-left -> left-to-right and vice-versa)
  ///
  var sheet1rtl = excel['Sheet1'].isRTL;
  //excel['Sheet1'].isRTL = false;
  print(
      'Sheet1: ((previous) isRTL: $sheet1rtl) ---> ((current) isRTL: ${excel['Sheet1'].isRTL})');
  var sheet1 = excel['Sheet1'];
  sheet1.cell(CellIndex.indexByString('A1')).value = 'Sheet1';

  excel.copy('Sheet1', 'newlyCopied');

  var sheet2 = excel['newlyCopied'];
  sheet2.cell(CellIndex.indexByString('A1')).value = 'newlyCopied';

  /* var sheet2rtl = excel['Sheet2'].isRTL;
  excel['Sheet2'].isRTL = true;
  print(
      'Sheet2: ((previous) isRTL: $sheet2rtl) ---> ((current) isRTL: ${excel['Sheet2'].isRTL})'); */

  ///
  ///
  /// declaring a cellStyle object
  ///
  ///
  /*  CellStyle cellStyle = CellStyle(
    bold: true,
    italic: true,
    textWrapping: TextWrapping.WrapText,
    fontFamily: getFontFamily(FontFamily.Comic_Sans_MS),
    rotation: 0,
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

  ///
  ///
  /// Iterating and changing values to desired type
  ///
  ///
  for (int row = 0; row < sheet.maxRows; row++) {
    sheet.row(row).forEach((Data? cell1) {
      if (cell1 != null) {
        cell1.value = ' My custom Value ';
      }
    });
  }

  excel.rename("mySheet", "myRenamedNewSheet");

  // fromSheet should exist in order to sucessfully copy the contents
  excel.copy('myRenamedNewSheet', 'toSheet');

  excel.rename('oldSheetName', 'newSheetName');

  excel.delete('Sheet1');

  excel.unLink('sheet1');

  sheet = excel['sheet'];

  /// appending rows and checking the time complexity of it
  Stopwatch stopwatch = Stopwatch()..start();
  List<List<String>> list = List.generate(
      9000, (index) => List.generate(20, (index1) => '$index $index1'));

  print('list creation executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  list.forEach((row) {
    sheet.appendRow(row);
  });
  print('appending executed in ${stopwatch.elapsed}');

  sheet.appendRow([8]);
  bool isSet = excel.setDefaultSheet(sheet.sheetName);
  // isSet is bool which tells that whether the setting of default sheet is successful or not.
  if (isSet) {
    print("${sheet.sheetName} is set to default sheet.");
  } else {
    print("Unable to set ${sheet.sheetName} to default sheet.");
  }

  var colIterableSheet = excel['ColumnIterables'];

  var colIterables = ['A', 'B', 'C', 'D', 'E'];
  int colIndex = 0;

  colIterables.forEach((colValue) {
    colIterableSheet.cell(CellIndex.indexByColumnRow(
      rowIndex: colIterableSheet.maxRows,
      columnIndex: colIndex,
    ))
      ..value = colValue;
  }); */

  // Saving the file

  String outputFile = "/Users/kawal/Desktop/git_projects/r.xlsx";

  //stopwatch.reset();
  List<int>? fileBytes = excel.save();
  //print('saving executed in ${stopwatch.elapsed}');
  if (fileBytes != null) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(fileBytes);
  }
}
