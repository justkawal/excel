import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/excel2.xlsx";
  var bytes = File(file).readAsBytesSync();
  var excel = Excel.createExcel();
  // or
  //var excel = Excel.decodeBytes(bytes, update: true);
  for (var table in excel.tables.keys) {
    print(table);
    print(excel.tables[table].maxCols);
    print(excel.tables[table].maxRows);
    for (var row in excel.tables[table].rows) {
      print("$row");
    }
  }

  // if sheet with name = Sheet24 does not exist then it will be automatically created.
  var sheet = 'Sheet24';

  excel.updateCell(sheet, CellIndex.indexByString("A1"), "Here Value of A1",
      backgroundColorHex: "#1AFF1A",
      horizontalAlign: HorizontalAlign.Center,
      verticalAlign: VerticalAlign.Center);

  excel.updateCell(
      sheet,
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
      "Here Value of C1",
      wrap: TextWrapping.WrapText);

  excel.updateCell(sheet, CellIndex.indexByString("A2"), "Here Value of A2",
      backgroundColorHex: "#1AFF1A", wrap: TextWrapping.Clip);

  excel.updateCell(sheet, CellIndex.indexByString("E5"), "E5",
      horizontalAlign: HorizontalAlign.Right, wrap: TextWrapping.Clip);

  excel.merge(
      sheet, CellIndex.indexByString("A1"), CellIndex.indexByString("E4"),
      customValue: "Now it is merged");

  excel.merge(
      sheet, CellIndex.indexByString("F1"), CellIndex.indexByString("F5"));

  excel.merge(
      sheet, CellIndex.indexByString("A5"), CellIndex.indexByString("E5"));

  //Remove row at index = 2
  excel.removeRow(sheet, 2);

  // Remove column at index = 2
  excel.removeColumn(sheet, 2);

  // Insert column at index = 2;
  excel.insertColumn(sheet, 2);

  // Insert row at index = 2;
  excel.insertRow(sheet, 2);

  excel.appendRow(sheet, ["bustin", "jiebr"]);

  int replacedCount = excel.findAndReplace(sheet, 'bustin', 'raman');
  print("Replaced Count:" + replacedCount.toString());

  excel.setDefaultSheet(sheet).then((isSet) {
    // isSet is bool which tells that whether the setting of default sheet is successful or not.
    if (isSet) {
      print("$sheet is set to default sheet.");
    } else {
      print("Unable to set $sheet to default sheet.");
    }
  });

  excel.getDefaultSheet().then((value) {
    print("Default Sheet:" + value.toString());
  });

  excel.insertRowIterables(sheet, ["A", "B", "C", "D", "E", "F", "G", "H"], 2,
      startingColumn: 3, overwriteMergedCells: false);

  excel.insertRowIterables(
      sheet, ["Insert", "ing", "in", "9th", "row", "as", "iterables"], 8,
      startingColumn: 13);

  // Check which cells are merged
  List<String> mergedCells = excel.getMergedCells(sheet);
  mergedCells.forEach((cells) {
    print("Merged:" + cells.toString());
  });

  // After removal of column and insertion of row merged - A1:E4 becomes merged - A1:D5
  // So we have to call un-merge at A1:D5
  if (mergedCells.contains("A1:D5")) {
    excel.unMerge(sheet, "A1:D5");
  }

  // Saving the file

  String outputFile = "/Users/kawal/Desktop/excel_example.xlsx";
  excel.encode().then((onValue) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
}
