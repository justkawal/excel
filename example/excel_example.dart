import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/form.xlsx";
  var bytes = File(file).readAsBytesSync();
  // var excel = Excel.createExcel();
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

  var sheet = excel['mySheet'];

  /// List of rows.
  sheet.rows;

  /// putting value = 'k' at A1 index.
  var cell = sheet.cell(CellIndex.indexByString("A1"));
  cell.value = "Heya How are you I am fine ok goood night";
  CellStyle _cellStyle = CellStyle();
  _cellStyle.isBold = true;
  _cellStyle.isItalic = true;
  cell.cellStyle = _cellStyle;

  /// putting value = 'k' at A1 index.
  var cell2 = sheet.cell(CellIndex.indexByString("A5"));
  cell2.value = "Heya How night";
  CellStyle _cellStyle2 = CellStyle();
  _cellStyle2.underline = Underline.Single;

  cell2.cellStyle = _cellStyle2;

  /// appending rows
  sheet.appendRow([8]);
  excel.setDefaultSheet(sheet.sheetName).then((isSet) {
    // isSet is bool which tells that whether the setting of default sheet is successful or not.
    if (isSet) {
      print("${sheet.sheetName} is set to default sheet.");
    } else {
      print("Unable to set ${sheet.sheetName} to default sheet.");
    }
  });
/* 
  /// coutn of rows
  sheet.maxRows;

  /// count of cols
  sheet.maxCols;

  sheet.clearRow(0);
  sheet.removeColumn(0);
  sheet.removeRow(0);
  sheet.insertColumn(0);

  // if sheet with name = Sheet24 does not exist then it will be automatically created.
  var sheet1 = 'Sheet24';

  //Remove row at index = 0
  excel.removeRow('sheet', 0);

  excel.updateCell(sheet1, CellIndex.indexByString("A1"), "Here Value of A1");

  excel.updateCell(
      sheet1,
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
      "Here Value of C1");

  excel.updateCell(sheet1, CellIndex.indexByString("A2"), "Here Value of A2");

  excel.updateCell(sheet1, CellIndex.indexByString("E5"), "E5");

  excel.merge(
      sheet1, CellIndex.indexByString("A1"), CellIndex.indexByString("E4"),
      customValue: "Now it is merged");

  excel.merge(
      sheet1, CellIndex.indexByString("F1"), CellIndex.indexByString("F5"));

  excel.merge(
      sheet1, CellIndex.indexByString("A5"), CellIndex.indexByString("E5"));

  //Remove row at index = 2
  excel.removeRow(sheet1, 2);

  // Remove column at index = 2
  excel.removeColumn(sheet1, 2);

  // Insert column at index = 2;
  excel.insertColumn(sheet1, 2);

  // Insert row at index = 2;
  excel.insertRow(sheet1, 2);

  excel.appendRow(sheet1, ["bustin", "jiebr"]);

  int replacedCount = excel.findAndReplace(sheet1, 'bustin', 'raman');
  print("Replaced Count:" + replacedCount.toString());


  excel.getDefaultSheet().then((value) {
    print("Default Sheet:" + value.toString());
  });

  excel.insertRowIterables(sheet1, ["A", "B", "C", "D", "E", "F", "G", "H"], 2,
      startingColumn: 3, overwriteMergedCells: false);

  excel.insertRowIterables(
      sheet1, ["Insert", "ing", "in", "9th", "row", "as", "iterables"], 8,
      startingColumn: 13);

  // Check which cells are merged
  List<String> mergedCells = excel.getMergedCells(sheet1);
  mergedCells.forEach((cells) {
    print("Merged:" + cells.toString());
  });

  // After removal of column and insertion of row merged - A1:E4 becomes merged - A1:D5
  // So we have to call un-merge at A1:D5
  if (mergedCells.contains("A1:D5")) {
    excel.unMerge(sheet1, "A1:D5");
  }

  /// copies the contents of
  excel['copiedInto'] = excel['Sheet24']; */

  // Saving the file

  String outputFile = "/Users/kawal/Desktop/error/form.xlsx";
  excel.encode().then((onValue) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
}
