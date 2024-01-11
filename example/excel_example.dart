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
    print(excel.tables[table]!.maxColumns);
    print(excel.tables[table]!.maxRows);
    for (var row in excel.tables[table]!.rows) {
      print("${row.map((e) => e?.value)}");
    }
  }

  ///
  /// Change sheet from rtl to ltr and vice-versa i.e. (right-to-left -> left-to-right and vice-versa)
  ///
  var sheet1rtl = excel['Sheet1'].isRTL;
  excel['Sheet1'].isRTL = false;
  print(
      'Sheet1: ((previous) isRTL: $sheet1rtl) ---> ((current) isRTL: ${excel['Sheet1'].isRTL})');

  var sheet2rtl = excel['Sheet2'].isRTL;
  excel['Sheet2'].isRTL = true;
  print(
      'Sheet2: ((previous) isRTL: $sheet2rtl) ---> ((current) isRTL: ${excel['Sheet2'].isRTL})');

  ///
  ///
  /// declaring a cellStyle object
  ///
  ///
  CellStyle cellStyle = CellStyle(
    bold: true,
    italic: true,
    textWrapping: TextWrapping.WrapText,
    fontFamily: getFontFamily(FontFamily.Comic_Sans_MS),
    rotation: 0,
  );

  var sheet = excel['mySheet'];

  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = TextCellValue("Heya How are you I am fine ok goood night");
  cell.cellStyle = cellStyle;

  var cell2 = sheet.cell(CellIndex.indexByString("E5"));
  cell2.value = TextCellValue("Heya How night");
  cell2.cellStyle = cellStyle;

  /// printing cell-type
  print("CellType: " +
      switch (cell.value) {
        null => 'empty',
        TextCellValue() => 'text',
        FormulaCellValue() => 'Formula',
        IntCellValue() => 'int',
        DoubleCellValue() => 'double',
        DateCellValue() => 'date',
        DateTimeCellValue() => 'date+time',
        TimeCellValue() => 'time',
        BoolCellValue() => 'bool',
      });

  ///
  ///
  /// Iterating and changing values to desired type
  ///
  ///
  for (int row = 0; row < sheet.maxRows; row++) {
    sheet.row(row).forEach((Data? cell1) {
      if (cell1 != null) {
        cell1.value = TextCellValue(' My custom Value ');
      }
    });
  }

  excel.rename("mySheet", "myRenamedNewSheet");

  var sheet1 = excel['Sheet1'];
  sheet1.cell(CellIndex.indexByString('A1')).value = TextCellValue('Sheet1');

  /// fromSheet should exist in order to sucessfully copy the contents
  excel.copy('Sheet1', 'newlyCopied');

  var sheet2 = excel['newlyCopied'];
  sheet2.cell(CellIndex.indexByString('A1')).value =
      TextCellValue('Newly Copied Sheet');

  /// renaming the sheet
  excel.rename('oldSheetName', 'newSheetName');

  /// deleting the sheet
  excel.delete('Sheet1');

  /// unlinking the sheet if any link function is used !!
  excel.unLink('sheet1');

  sheet = excel['sheet'];

  /// appending rows and checking the time complexity of it
  Stopwatch stopwatch = Stopwatch()..start();
  List<List<TextCellValue>> list = List.generate(
    9000,
    (index) => List.generate(20, (index1) => TextCellValue('$index $index1')),
  );

  print('list creation executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  list.forEach((row) {
    sheet.appendRow(row);
  });
  print('appending executed in ${stopwatch.elapsed}');

  sheet.appendRow([
    IntCellValue(8),
    DoubleCellValue(999.62221),
    DateCellValue(
      year: DateTime.now().year,
      month: DateTime.now().month,
      day: DateTime.now().day,
    ),
    DateTimeCellValue.fromDateTime(DateTime.now()),
  ]);

  bool isSet = excel.setDefaultSheet(sheet.sheetName);
  // isSet is bool which tells that whether the setting of default sheet is successful or not.
  if (isSet) {
    print("${sheet.sheetName} is set to default sheet.");
  } else {
    print("Unable to set ${sheet.sheetName} to default sheet.");
  }

  var columnIterableSheet = excel['ColumnIterables'];

  var columnIterables = ['A', 'B', 'C', 'D', 'E'];
  int columnIndex = 0;

  columnIterables.forEach((columnValue) {
    columnIterableSheet.cell(CellIndex.indexByColumnRow(
      rowIndex: columnIterableSheet.maxRows,
      columnIndex: columnIndex,
    ))
      ..value = TextCellValue(columnValue);
  });

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
