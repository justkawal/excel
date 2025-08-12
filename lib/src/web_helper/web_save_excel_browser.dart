import 'dart:js_interop';
import 'dart:typed_data';

import 'package:cross_file/cross_file.dart';
import 'package:web/web.dart';

// A wrapper to save the excel file in browser
class SavingHelper {
  static List<int>? saveFile(List<int>? val, String fileName) {
    if (val == null) {
      return null;
    }

    final blob = Blob(JSArray.from(Uint8List.fromList(val).toJS));
    final url = URL.createObjectURL(blob);
    final anchor = HTMLAnchorElement()
      ..href = url
      ..download = '$fileName';

    document.body?.append(anchor);

    // download the file
    anchor.click();

    // cleanup
    anchor.remove();
    URL.revokeObjectURL(url);
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
