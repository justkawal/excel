import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  var file = "/Users/kawal/Desktop/excel2.xlsx";
  var bytes = File(file).readAsBytesSync();
  var updater = Excel.createExcel(); // .decodeBytes(bytes, update: true);
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

  int getPosition(String val) {
    List<String> spannedCells = updater.getSpannedItems(sheet);
    return spannedCells.indexOf(val);
  }

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

  updater.removeColumn(sheet, 2);
  updater.insertRow(sheet, 2);

  // updater.unMerge(sheet, getPosition("A1:E4"));

  // Saving the file

  String outputFile = "/Users/kawal/Desktop/excel2copy.xlsx";
  updater.encode().then((onValue) {
    File(join(outputFile))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
}
