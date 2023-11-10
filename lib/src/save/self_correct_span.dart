part of excel;

///Self correct the spanning of rows and columns by checking their cross-sectional relationship between if exists.
_selfCorrectSpanMap(Excel _excel) {
  _excel._mergeChangeLook.forEach((String key) {
    if (_excel._sheetMap[key] != null &&
        _excel._sheetMap[key]!._spanList.isNotEmpty) {
      List<_Span?> spanList =
          List<_Span?>.from(_excel._sheetMap[key]!._spanList);

      for (int i = 0; i < spanList.length; i++) {
        _Span? checkerPos = spanList[i];
        if (checkerPos == null) {
          continue;
        }
        int startRow = checkerPos.rowSpanStart,
            startColumn = checkerPos.columnSpanStart,
            endRow = checkerPos.rowSpanEnd,
            endColumn = checkerPos.columnSpanEnd;

        for (int j = i + 1; j < spanList.length; j++) {
          _Span? spanObj = spanList[j];
          if (spanObj == null) {
            continue;
          }

          final locationChange = _isLocationChangeRequired(
              startColumn, startRow, endColumn, endRow, spanObj);
          if (locationChange.$1) {
            startColumn = locationChange.$2.$1;
            startRow = locationChange.$2.$2;
            endColumn = locationChange.$2.$3;
            endRow = locationChange.$2.$4;
            spanList[j] = null;
          } else {
            final locationChange2 = _isLocationChangeRequired(
                spanObj.columnSpanStart,
                spanObj.rowSpanStart,
                spanObj.columnSpanEnd,
                spanObj.rowSpanEnd,
                checkerPos);

            if (locationChange2.$1) {
              startColumn = locationChange2.$2.$1;
              startRow = locationChange2.$2.$2;
              endColumn = locationChange2.$2.$3;
              endRow = locationChange2.$2.$4;
              spanList[j] = null;
            }
          }
        }
        _Span spanObj1 = _Span(
          rowSpanStart: startRow,
          columnSpanStart: startColumn,
          rowSpanEnd: endRow,
          columnSpanEnd: endColumn,
        );
        spanList[i] = spanObj1;
      }
      _excel._sheetMap[key]!._spanList = List<_Span?>.from(spanList);
      _excel._sheetMap[key]!._cleanUpSpanMap();
    }
  });
}
