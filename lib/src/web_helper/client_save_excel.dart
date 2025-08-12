import 'dart:typed_data';

import 'package:cross_file/cross_file.dart';

class SavingHelper {
// A wrapper to save the excel file in client
  static List<int>? saveFile(List<int>? val, String fileName) {
    return val;
  }

  static XFile? generateXFile(List<int>? val, String fileName) {
    if (val == null) return null;
    final bytes = Uint8List.fromList(val);
    return XFile.fromData(
      bytes,
      name: fileName,
      length: bytes.length,
      lastModified: DateTime.now(),
      mimeType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      path: fileName,
    );
  }
}
