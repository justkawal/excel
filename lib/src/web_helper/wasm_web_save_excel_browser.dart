import 'dart:js_interop';
import 'dart:typed_data';

import 'package:web/web.dart';

class SavingHelper {
  static List<int>? saveFile(List<int>? val, String fileName) {
    if (val == null) {
      return null;
    }

    // Create Uint8List
    final uint8List = Uint8List.fromList(val);

    // Convert to JS object properly
    final jsData = uint8List.toJS;

    // Create blob
    final blob = Blob(
        [jsData].toJS,
        BlobPropertyBag(
            type:
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'));

    final url = URL.createObjectURL(blob);
    final anchor = HTMLAnchorElement()
      ..href = url
      ..download = fileName;

    document.body?.append(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);

    return val;
  }
}
