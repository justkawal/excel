import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/excel2.xlsx";
  var bytes = File(file).readAsBytesSync();
  var updater = Excel.createExcel();
  // or
  //var updater = Excel.decodeBytes(bytes, update: true);
  for (var table in updater.tables.keys) {
    print(table);
    print(updater.tables[table].maxCols);
    print(updater.tables[table].maxRows);
    for (var row in updater.tables[table].rows) {
      print("$row");
    }
  }

  // if sheet with name = Sheet24 does not exist then it will be automatically created.
  var sheet = 'Sheet24';

  updater.updateCell(sheet, CellIndex.indexByString("A1"), "Here Value of A1",
      backgroundColorHex: "#1AFF1A", verticalAlign: VerticalAlign.Bottom);

  updater.updateCell(
      sheet,
      CellIndex.indexByColumnRow(columnIndex: 2, rowIndex: 0),
      "Here Value of C1",
      wrap: TextWrapping.WrapText);

  updater.updateCell(sheet, CellIndex.indexByString("A2"), "Here Value of A2",
      backgroundColorHex: "#1AFF1A", wrap: TextWrapping.Clip);

  updater.updateCell(sheet, CellIndex.indexByString("E5"), "E5",
      horizontalAlign: HorizontalAlign.Right, wrap: TextWrapping.Clip);

  updater.merge(
      sheet, CellIndex.indexByString("A1"), CellIndex.indexByString("E4"),
      customValue: "Now it is merged");

  // Insert column at index = 17;
  updater.insertColumn(sheet, 17);

  // Remove column at index = 2
  updater.removeColumn(sheet, 2);

  // Insert row at index = 2;
  updater.insertRow(sheet, 2);

  // Remove row at index = 17
  updater.removeRow(sheet, 2);

  updater.setDefaultSheet(sheet).then((isSet) {
    // isSet is bool which tells that whether the setting of default sheet is successful or not.
    if (isSet) {
      print("$sheet is set to default sheet.");
    } else {
      print("Unable to set $sheet to default sheet.");
    }
  });

  updater.getDefaultSheet().then((value) {
    print("Default Sheet:" + value.toString());
  });

  // Check which cells are merged
  List<String> mergedCells = updater.getMergedCells(sheet);
  mergedCells.forEach((cells) {
    print("Merged:" + cells.toString());
  });

  // After removal of column and insertion of row merged - A1:E4 becomes merged - A1:D5
  // So we have to call un-merge at A1:D5
  if (mergedCells.contains("A1:D5")) {
    updater.unMerge(sheet, "A1:D5");
  }

  // Saving the file

  String outputFile = "/Users/kawal/Desktop/excel2copy.xlsx";
  updater.encode().then((onValue) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
}
