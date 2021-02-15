import 'dart:io';
import 'package:path/path.dart';
import 'package:excel/excel.dart';

void main(List<String> args) {
  Stopwatch stopwatch = new Stopwatch()..start();

  Excel excel = Excel.createExcel();
  Sheet sh = excel['Sheet1'];
  for (int i = 0; i < 8; i++) {
    sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).value =
        'Col $i';
    sh.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: i)).cellStyle =
        CellStyle(bold: true);
  }
  for (int row = 1; row < 9000; row++) {
    for (int col = 0; col < 8; col++) {
      sh
          .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row))
          .value = 'value ${row}_$col';
    }
  }
  print('Generating executed in ${stopwatch.elapsed}');
  stopwatch.reset();
  excel.encode().then((onValue) {
    print('Encoding executed in ${stopwatch.elapsed}');
    stopwatch.reset();
    File(join("/Users/kawal/Desktop/r1.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
    print('Downloaded executed in ${stopwatch.elapsed}');
  });
}
