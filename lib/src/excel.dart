part of excel;

const String _relationships =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const _spreasheetXlsx = 'xlsx';

enum TextWrapping { WrapText, Clip }
enum VerticalAlign { Top, Center, Bottom }
enum HorizontalAlign { Left, Center, Right }

// Normalize line
String _normalizeNewLine(String text) {
  return text.replaceAll('\r\n', '\n');
}

/// Returns the coordinates from a cell name.
/// "A2" returns [2, 1] and the "B3" return [3, 2].
List<int> cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart =
      utf8.decode(letters.where((rune) => rune > 0).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);

  return [
    int.parse(numericsPart) - 1,
    lettersToNumeric(lettersPart) - 1
  ]; // [x , y]
}

Excel _newExcel(Archive archive, bool update) {
  // Lookup at file format
  var format;

  // Try OpenDocument format
  var mimetype = archive.findFile('mimetype');
  if (mimetype == null) {
    var xl = archive.findFile('xl/workbook.xml');
    if (xl != null) {
      format = _spreasheetXlsx;
    }
  }

  switch (format) {
    case _spreasheetXlsx:
      return XlsxDecoder(archive, update: update);
    default:
      throw UnsupportedError('Excel format unsupported');
  }
}

List<String> _noCompression = <String>[
  'mimetype',
  'Thumbnails/thumbnail.png',
];

/// Convert a number to character based column
String numericToLetters(int number) {
  var letters = '';

  while (number != 0) {
    // Set remainder from 1..26
    var remainder = number % 26;

    if (remainder == 0) {
      remainder = 26;
    }

    // Convert the remainder to a character.
    var letter = String.fromCharCode(65 + remainder - 1);

    // Accumulate the column letters, right to left.
    letters = letter + letters;

    // Get the next order of magnitude.
    number = (number - 1) ~/ 26;
  }
  return letters;
}

/// Decode a excel file.
abstract class Excel {
  bool _update, _colorChanges, _mergeChanges;
  Archive _archive;
  Map<String, XmlNode> _sheets;
  Map<String, XmlDocument> _xmlFiles;
  Map<String, String> _xmlSheetId;
  Map<String, ArchiveFile> _archiveFiles;
  Map<String, String> _worksheetTargets;
  Map<String, Map<String, CellStyle>> _cellStyleOther;
  Map<String, Map<String, int>> _cellStyleReferenced;
  Map<String, List<String>> _spannedItems;
  Map<String, DataTable> _tables;
  List<CellStyle> _cellStyleList, _innerCellStyle;
  Map<String, List<_Span>> _spanMap;
  List<String> _sharedStrings,
      _rId,
      _fontColorHex,
      _patternFill,
      _mergeChangeLook;
  List<int> _numFormats;
  String _stylesTarget, _sharedStringsTarget;

  /// Filename extension
  String get extension;

  /// Tables contained in excel file indexed by their names
  Map<String, DataTable> get tables => _tables;

  Excel();

  factory Excel.createExcel() {
    String newSheet =
        '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';
    return Excel.decodeBytes(Base64Decoder().convert(newSheet), update: true);
  }

  factory Excel.decodeBytes(List<int> data,
      {bool update = false, bool verify = false}) {
    var archive = ZipDecoder().decodeBytes(data, verify: verify);
    return _newExcel(archive, update);
  }

  factory Excel.decodeBuffer(InputStream input,
      {bool update = false, bool verify = false}) {
    var archive = ZipDecoder().decodeBuffer(input, verify: verify);
    return _newExcel(archive, update);
  }

  _damagedExcel({String text}) {
    String t = '\nDamaged Excel file:';
    if (text != null) {
      t += ' $text';
    }
    throw ArgumentError(t + '\n');
  }

  int _getAvailableRid() {
    _rId.sort((a, b) =>
        int.parse(a.substring(3)).compareTo(int.parse(b.substring(3))));

    List<String> got = List<String>.from(_rId.last.split(''));
    got.removeWhere((item) => !'0123456789'.split('').contains(item));
    return int.parse(got.join().toString()) + 1;
  }

  /// Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  /// Add the sheet details in the workbook.xml. as well as in the workbook.xml.rels
  /// Then add the sheet physically into the [_xmlFiles] so as to get it into the archieve.
  /// Also add it into the [_sheets] and [_tables] map so as to allow the editing.
  _createSheet(String newSheet) {
    List<XmlNode> list =
        _xmlFiles['xl/workbook.xml'].findAllElements('sheets').first.children;
    if (list.isEmpty) {
      throw ArgumentError('');
    }
    int _sheetId = -1;
    List<int> sheetIdList = List<int>();

    _xmlFiles['xl/workbook.xml']
        .findAllElements('sheet')
        .forEach((sheetIdNode) {
      var sheetId = sheetIdNode.getAttribute('sheetId');
      if (sheetId != null) {
        int t = int.parse(sheetId.toString());
        if (!sheetIdList.contains(t)) {
          sheetIdList.add(t);
        }
      } else {
        _damagedExcel(text: 'Corrupted Sheet Indexing');
      }
    });

    sheetIdList.sort();

    for (int i = 0; i < sheetIdList.length - 1; i++) {
      if ((sheetIdList[i] + 1) != sheetIdList[i + 1]) {
        _sheetId = (sheetIdList[i] + 1);
      }
    }
    if (_sheetId == -1) {
      if (sheetIdList.isEmpty) {
        _sheetId = 1;
      } else {
        _sheetId = sheetIdList.length;
      }
    }

    int sheetNumber = _sheetId;
    int ridNumber = _getAvailableRid();

    _xmlFiles['xl/_rels/workbook.xml.rels']
        .findAllElements('Relationships')
        .first
        .children
        .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName('Type'), '$_relationships/worksheet'),
          XmlAttribute(
              XmlName('Target'), 'worksheets/sheet${sheetNumber + 1}.xml'),
        ]));

    _xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children
        .add(XmlElement(
          XmlName('sheet'),
          <XmlAttribute>[
            XmlAttribute(XmlName('state'), 'visible'),
            XmlAttribute(XmlName('name'), newSheet),
            XmlAttribute(XmlName('sheetId'), '${sheetNumber + 1}'),
            XmlAttribute(XmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] =
        'worksheets/sheet${sheetNumber + 1}.xml';

    var content = utf8.encode(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView tabSelected=\"1\" workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

    _archive.addFile(ArchiveFile(
        'xl/worksheets/sheet${sheetNumber + 1}.xml', content.length, content));
    var _newSheet = _archive.findFile('xl/$_sharedStringsTarget');

    _newSheet.decompress();
    var document = parse(utf8.decode(_newSheet.content));
    if (_xmlFiles != null) {
      _xmlFiles['xl/worksheets/sheet${sheetNumber + 1}.xml'] = document;
    }

    _xmlFiles['[Content_Types].xml']
        .findAllElements('Types')
        .first
        .children
        .add(XmlElement(
          XmlName('Override'),
          <XmlAttribute>[
            XmlAttribute(XmlName('ContentType'),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            XmlAttribute(XmlName('PartName'),
                '/xl/worksheets/sheet${sheetNumber + 1}.xml'),
          ],
        ));
    _parseTable(_xmlFiles['xl/workbook.xml'].findAllElements('sheet').last);
  }

  // Reading the styles from the excel file.
  _parseStyles(String _stylesTarget) {
    var styles = _archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = parse(utf8.decode(styles.content));
      if (_xmlFiles != null) {
        _xmlFiles['xl/$_stylesTarget'] = document;
      }
      _fontColorHex = List<String>();
      _patternFill = List<String>();
      _cellStyleList = List<CellStyle>();
      int fontIndex = 0;
      document
          .findAllElements('font')
          .first
          .findElements('color')
          .forEach((child) {
        var colorHex = child.getAttribute('rgb');
        if (colorHex != null && !_fontColorHex.contains(colorHex.toString())) {
          _fontColorHex.add(colorHex.toString());
          fontIndex = 1;
        } else if (fontIndex == 0 && !_fontColorHex.contains("FF000000")) {
          _fontColorHex.add("FF000000");
        }
      });
      document.findAllElements('patternFill').forEach((node) {
        String patternType = node.getAttribute('patternType').toString(), rgb;
        if (node.children.isNotEmpty) {
          node.findElements('fgColor').forEach((child) {
            rgb = node.getAttribute('rgb').toString();
            _patternFill.add(rgb);
          });
        } else {
          _patternFill.add(patternType);
        }
      });

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          _numFormats.add(_getFontIndex(node, 'numFmtId'));

          String fontColor = "FF000000", backgroundColor = "none";
          HorizontalAlign horizontalAlign = HorizontalAlign.Left;
          VerticalAlign verticalAlign = VerticalAlign.Bottom;
          TextWrapping textWrapping;
          int fontId = _getFontIndex(node, 'fontId');
          if (fontId < _fontColorHex.length) {
            fontColor = _fontColorHex[fontId];
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _patternFill.length) {
            backgroundColor = _patternFill[fillId];
          }

          if (node.children.isNotEmpty) {
            node.findElements('alignment').forEach((child) {
              if (_getFontIndex(child, 'wrapText') == 1) {
                textWrapping = TextWrapping.WrapText;
              } else if (_getFontIndex(child, 'shrinkToFit') == 1) {
                textWrapping = TextWrapping.Clip;
              }

              var vertical = node.getAttribute('vertical');
              if (vertical != null) {
                if (vertical.toString() == 'top') {
                  verticalAlign = VerticalAlign.Top;
                } else if (vertical.toString() == 'center') {
                  verticalAlign = VerticalAlign.Center;
                }
              }

              var horizontal = node.getAttribute('horizontal');
              if (horizontal != null) {
                if (horizontal.toString() == 'center') {
                  horizontalAlign = HorizontalAlign.Center;
                } else if (horizontal.toString() == 'right') {
                  horizontalAlign = HorizontalAlign.Right;
                }
              }
            });
          }

          CellStyle cellStyle = CellStyle(
              fontColorHex: fontColor,
              backgroundColorHex: backgroundColor,
              horizontalAlign: horizontalAlign,
              verticalAlign: verticalAlign,
              textWrapping: textWrapping);

          _cellStyleList.add(cellStyle);
        });
      });
    } else {
      _damagedExcel(text: 'styles');
    }
  }

  int _getFontIndex(var node, String text) {
    int applyFontInt = 0;
    var applyFont = node.getAttribute(text);
    if (applyFont != null) {
      try {
        applyFontInt = int.parse(applyFont.toString());
      } catch (_) {}
    }
    return applyFontInt;
  }

  /// Encode bytes after update
  Future<List> encode() async {
    if (!_update) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }

    if (_colorChanges) {
      _processStylesFile();
    }
    _setSheetElements();
    _setSharedStrings();

    if (_mergeChanges) {
      _setMerge();
    }

    for (var xmlFile in _xmlFiles.keys) {
      var xml = _xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_archive));
  }

  /// It returns the name of the default sheet.
  Future<String> getDefaultSheet() async {
    XmlElement _sheet =
        _xmlFiles['xl/workbook.xml'].findAllElements('sheet').first;

    if (_sheet != null) {
      var defaultSheet = _sheet.getAttribute('name');
      if (defaultSheet != null) {
        return defaultSheet.toString();
      } else {
        _damagedExcel(
            text: 'Excel sheet corrupted!! Try creating new excel file.');
      }
    }
    return null;
  }

  /// It returns to true if the passed sheetName is set to default sheet otherwise returns false
  Future<bool> setDefaultSheet(String sheetName) async {
    int position = -1;
    List<XmlElement> sheetList =
        _xmlFiles['xl/workbook.xml'].findAllElements('sheet').toList();
    XmlElement elementFound;

    for (int i = 0; i < sheetList.length; i++) {
      var _sheetName = sheetList[i].getAttribute('name');
      if (_sheetName != null && _sheetName.toString() == sheetName) {
        elementFound = sheetList[i];
        position = i;
        break;
      }
    }

    if (position == -1) {
      return false;
    }
    if (position == 0) {
      return true;
    }

    _xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children
        .removeAt(position);

    _xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children
        .insert(0, elementFound);

    String expectedSheet = await getDefaultSheet();

    return expectedSheet == sheetName;
  }

  Archive _cloneArchive(Archive archive) {
    var clone = Archive();
    archive.files.forEach((file) {
      if (file.isFile) {
        ArchiveFile copy;
        if (_archiveFiles.containsKey(file.name)) {
          copy = _archiveFiles[file.name];
        } else {
          var content = (file.content as Uint8List).toList();
          var compress = !_noCompression.contains(file.name);
          copy = ArchiveFile(file.name, content.length, content)
            ..compress = compress;
        }
        clone.addFile(copy);
      }
    });
    return clone;
  }

  _normalizeTable(DataTable table) {
    if (table._maxRows == 0) {
      table._rows.clear();
    } else if (table._maxRows < table._rows.length) {
      table._rows.removeRange(table._maxRows, table._rows.length);
    }

    for (var row = 0; row < table._rows.length; row++) {
      if (table._maxCols == 0) {
        table._rows[row].clear();
      } else if (table._maxCols < table._rows[row].length) {
        table._rows[row].removeRange(table._maxCols, table._rows[row].length);
      } else if (table._maxCols > table._rows[row].length) {
        var repeat = table._maxCols - table._rows[row].length;
        for (var index = 0; index < repeat; index++) {
          table._rows[row].add(null);
        }
      }
    }
  }

  int _checkPosition(List<CellStyle> list, CellStyle cellStyle) {
    for (int i = 0; i < list.length; i++) {
      if (list[i] == cellStyle) {
        return i;
      }
    }
    return -1;
  }

  /// Writing Font Color in [xl/styles.xml] from the Cells of the sheets.

  _processStylesFile() {
    _innerCellStyle = List<CellStyle>();
    List<String> innerPatternFill = List<String>(),
        innerFontColor = List<String>();

    _cellStyleOther.keys.toList().forEach((otherSheet) {
      _cellStyleOther[otherSheet].forEach((String _, CellStyle cellStyleOther) {
        int pos = _checkPosition(_innerCellStyle, cellStyleOther);
        if (pos == -1) {
          _innerCellStyle.add(cellStyleOther);
        }
      });
    });

    _innerCellStyle.forEach((cellStyle) {
      String fontColor = cellStyle.getFontColorHex,
          backgroundColor = cellStyle.getBackgroundColorHex;

      if (!_fontColorHex.contains(fontColor) &&
          !innerFontColor.contains(fontColor)) {
        innerFontColor.add(fontColor);
      }
      if (!_patternFill.contains(backgroundColor) &&
          !innerPatternFill.contains(backgroundColor)) {
        innerPatternFill.add(backgroundColor);
      }
    });

    XmlElement fonts =
        _xmlFiles['xl/styles.xml'].findAllElements('fonts').first;

    var fontAttribute = fonts.getAttributeNode('count');
    if (fontAttribute != null) {
      fontAttribute.value = '${_fontColorHex.length + innerFontColor.length}';
    } else {
      fonts.attributes.add(XmlAttribute(
          XmlName('count'), '${_fontColorHex.length + innerFontColor.length}'));
    }

    innerFontColor.forEach((colorValue) =>
        fonts.children.add(XmlElement(XmlName('font'), [], [
          XmlElement(
              XmlName('color'), [XmlAttribute(XmlName('rgb'), colorValue)], [])
        ])));

    XmlElement fills =
        _xmlFiles['xl/styles.xml'].findAllElements('fills').first;

    var fillAttribute = fills.getAttributeNode('count');

    if (fillAttribute != null) {
      fillAttribute.value = '${_patternFill.length + innerPatternFill.length}';
    } else {
      fills.attributes.add(XmlAttribute(XmlName('count'),
          '${_patternFill.length + innerPatternFill.length}'));
    }

    innerPatternFill.forEach((color) {
      if (color.length >= 2) {
        if (color.substring(0, 2).toUpperCase() == 'FF') {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'), [
              XmlAttribute(XmlName('patternType'), 'solid')
            ], [
              XmlElement(XmlName('fgColor'),
                  [XmlAttribute(XmlName('rgb'), color)], []),
              XmlElement(
                  XmlName('bgColor'), [XmlAttribute(XmlName('rgb'), color)], [])
            ])
          ]));
        } else if (color == "none" ||
            color == "gray125" ||
            color == "lightGray") {
          fills.children.add(XmlElement(XmlName('fill'), [], [
            XmlElement(XmlName('patternFill'),
                [XmlAttribute(XmlName('patternType'), color)], [])
          ]));
        }
      } else {
        _damagedExcel(text: "Corrupted Styles Found");
      }
    });

    XmlElement celx =
        _xmlFiles['xl/styles.xml'].findAllElements('cellXfs').first;
    var cellAttribute = celx.getAttributeNode('count');

    if (cellAttribute != null) {
      cellAttribute.value = '${_cellStyleList.length + _innerCellStyle.length}';
    } else {
      celx.attributes.add(XmlAttribute(XmlName('count'),
          '${_cellStyleList.length + _innerCellStyle.length}'));
    }

    _innerCellStyle.forEach((cellStyle) {
      String backgroundColor = cellStyle.getBackgroundColorHex,
          fontColor = cellStyle.getFontColorHex;

      HorizontalAlign horizontalALign = cellStyle.getHorizontalAlignment;
      VerticalAlign verticalAlign = cellStyle.getVericalAlignment;
      TextWrapping textWrapping = cellStyle.getTextWrapping;
      int backgroundIndex = innerPatternFill.indexOf(backgroundColor),
          fontIndex = innerFontColor.indexOf(fontColor);

      var attributes = <XmlAttribute>[
        XmlAttribute(XmlName('borderId'), '0'),
        XmlAttribute(XmlName('fillId'),
            '${backgroundIndex == -1 ? 0 : backgroundIndex + _patternFill.length}'),
        XmlAttribute(XmlName('fontId'),
            '${fontIndex == -1 ? 0 : fontIndex + _fontColorHex.length}'),
        XmlAttribute(XmlName('numFmtId'), '0'),
        XmlAttribute(XmlName('xfId'), '0'),
      ];

      if ((_patternFill.contains(backgroundColor) ||
              innerPatternFill.contains(backgroundColor)) &&
          backgroundColor != "none" &&
          backgroundColor != "gray125" &&
          backgroundColor.toLowerCase() != "lightgray") {
        attributes.add(XmlAttribute(XmlName('applyFill'), '1'));
      }

      if ((_fontColorHex.contains(fontColor) ||
          innerFontColor.contains(fontColor))) {
        attributes.add(XmlAttribute(XmlName('applyFont'), '1'));
      }

      var children = <XmlElement>[];

      if (horizontalALign != HorizontalAlign.Left ||
          textWrapping != null ||
          verticalAlign != VerticalAlign.Bottom) {
        attributes.add(XmlAttribute(XmlName('applyAlignment'), '1'));
        var childAttributes = <XmlAttribute>[];

        if (textWrapping != null) {
          childAttributes.add(XmlAttribute(
              XmlName(textWrapping == TextWrapping.Clip
                  ? 'shrinkToFit'
                  : 'wrapText'),
              '1'));
        }

        if (verticalAlign != VerticalAlign.Bottom) {
          String ver = verticalAlign == VerticalAlign.Top ? 'top' : 'center';
          childAttributes.add(XmlAttribute(XmlName('vertical'), '$ver'));
        }

        if (horizontalALign != HorizontalAlign.Left) {
          String hor =
              horizontalALign == HorizontalAlign.Right ? 'right' : 'center';
          childAttributes.add(XmlAttribute(XmlName('horizontal'), '$hor'));
        }

        children.add(XmlElement(XmlName('alignment'), childAttributes, []));
      }

      celx.children.add(XmlElement(XmlName('xf'), attributes, children));
    });
  }

  /// Writing the value of excel cells into the separate
  /// sharedStrings file so as to minimize the size of excel files.
  _setSharedStrings() {
    String count = _sharedStrings.length.toString();
    List uniqueList = _sharedStrings.toSet().toList();
    String uniqueCount = uniqueList.length.toString();

    XmlElement shareString =
        _xmlFiles['xl/$_sharedStringsTarget'].findAllElements('sst').first;

    [
      ['count', count],
      ['uniqueCount', uniqueCount]
    ].forEach((value) {
      if (shareString.getAttributeNode(value[0]) == null) {
        shareString.attributes.add(XmlAttribute(XmlName(value[0]), value[1]));
      } else {
        shareString.getAttributeNode(value[0]).value = value[1];
      }
    });

    shareString.children.clear();

    _sharedStrings.forEach((string) {
      shareString.children.add(XmlElement(XmlName('si'), [], [
        XmlElement(XmlName('t'), [], [XmlText(string)])
      ]));
    });
  }

  ///Self correct the spanning of rows and columns by checking their cross-sectional relationship between if exists.
  _selfCorrectSpanMap() {
    _mergeChangeLook.forEach((key) {
      if (_spanMap.containsKey(key) && _tables.containsKey(key)) {
        for (int i = 0; i < _spanMap[key].length; i++) {
          if (_spanMap[key][i] != null) {
            _Span checkerPos = _spanMap[key][i];
            int startRow = checkerPos.rowSpanStart,
                startColumn = checkerPos.columnSpanStart,
                endRow = checkerPos.rowSpanEnd,
                endColumn = checkerPos.columnSpanEnd;

            for (int j = i + 1; j < _spanMap[key].length; j++) {
              if (_spanMap[key][j] != null) {
                _Span spanObj = _spanMap[key][j];

                Map gotMap = _isLocationChangeRequired(
                    key, startColumn, startRow, endColumn, endRow, spanObj);
                List<int> gotPosition = gotMap["gotPosition"],
                    changeValue = gotMap["changeValue"][0];

                if (changeValue == 1) {
                  startColumn = gotPosition[0];
                  startRow = gotPosition[1];
                  endColumn = gotPosition[2];
                  endRow = gotPosition[3];
                  _spanMap[key][j] = null;
                }
              }
            }
            _Span spanObj1 = _Span();
            spanObj1._start = [startRow, startColumn];
            spanObj1._end = [endRow, endColumn];
            _spanMap[key][i] = spanObj1;
          }
        }
      }
    });

    _mergeChangeLook.forEach((key) {
      if (_spanMap.containsKey(key)) {
        List<_Span> spanObjList = _spanMap[key];
        if (_tables.containsKey(key)) {
          List spanList = List<String>();
          spanObjList.forEach((value) {
            _Span spanObj = value;
            String rC = _getSpanCellId(
                spanObj.columnSpanStart,
                spanObj.rowSpanStart,
                spanObj.columnSpanEnd,
                spanObj.rowSpanEnd);
            if (!spanList.contains(rC)) {
              spanList.add(rC);
            }
          });
          _spannedItems[key] = spanList;
        }
      }
    });
  }

  /// Writing the merged cells information into the excel properties files.
  _setMerge() {
    _selfCorrectSpanMap();
    _mergeChangeLook.forEach((s) {
      if (_spannedItems.containsKey(s) &&
          _spanMap.containsKey(s) &&
          _xmlSheetId.containsKey(s) &&
          _xmlFiles.containsKey(_xmlSheetId[s])) {
        Iterable<XmlElement> iterMergeElement =
            _xmlFiles[_xmlSheetId[s]].findAllElements('mergeCells');
        XmlElement mergeElement;
        if (iterMergeElement.isNotEmpty) {
          mergeElement = iterMergeElement.first;
        } else {
          if (_xmlFiles[_xmlSheetId[s]].findAllElements('worksheet').length >
              0) {
            int index = _xmlFiles[_xmlSheetId[s]]
                .findAllElements('worksheet')
                .first
                .children
                .indexOf(_xmlFiles[_xmlSheetId[s]]
                    .findAllElements("sheetData")
                    .first);
            if (index == -1) {
              _damagedExcel();
            }
            _xmlFiles[_xmlSheetId[s]]
                .findAllElements('worksheet')
                .first
                .children
                .insert(
                    index + 1,
                    XmlElement(XmlName('mergeCells'),
                        [XmlAttribute(XmlName('count'), '0')]));

            mergeElement =
                _xmlFiles[_xmlSheetId[s]].findAllElements('mergeCells').first;
          } else {
            _damagedExcel();
          }
        }

        [
          ['count', _spannedItems[s].length.toString()],
        ].forEach((value) {
          if (mergeElement.getAttributeNode(value[0]) == null) {
            mergeElement.attributes
                .add(XmlAttribute(XmlName(value[0]), value[1]));
          } else {
            mergeElement.getAttributeNode(value[0]).value = value[1];
          }
        });

        mergeElement.children.clear();

        _spannedItems[s].forEach((value) {
          mergeElement.children.add(XmlElement(XmlName('mergeCell'),
              [XmlAttribute(XmlName('ref'), '$value')], []));
        });
      }
    });
  }

  /// Writing cell contained text into the excel sheet files.
  _setSheetElements() {
    _sharedStrings = List<String>();
    _tables.forEach((sheet, table) {
      for (int rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
        for (int columnIndex = 0;
            columnIndex < table.rows[rowIndex].length;
            columnIndex++) {
          if (table.rows[rowIndex][columnIndex] != null) {
            var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
            _updateCell(sheet, foundRow, columnIndex, rowIndex,
                table.rows[rowIndex][columnIndex]);
          }
        }
      }
    });
  }

  _checkSheetArguments(String sheet) {
    if (!_update) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }
    if (!_sheets.containsKey(sheet)) {
      _createSheet(sheet);
    }
  }

  /// Check if columnIndex is not out of Excel Column limits.
  _checkSheetMaxCol(String sheet, int colIndex) {
    if ((_tables[sheet]._maxCols >= 16384) || colIndex >= 16384) {
      throw ArgumentError('Reached Max (16384) or (XFD) columns value.');
    }
  }

  /// Check if rowIndex is not out of Excel Row limits.
  _checkSheetMaxRow(String sheet, int rowIndex) {
    if ((_tables[sheet]._maxRows >= 1048576) || rowIndex >= 1048576) {
      throw ArgumentError('Reached Max (1048576) rows value.');
    }
  }

  /// Insert column in [sheet] at position [columnIndex]
  insertColumn(String sheet, int columnIndex) {
    _checkSheetArguments(sheet);
    _checkSheetMaxCol(sheet, columnIndex);
    if (columnIndex < 0) {
      throw RangeError.range(columnIndex, 0, _tables[sheet]._maxCols);
    }

    bool updateSpanCell = false;

    if (_spanMap.containsKey(sheet)) {
      List spannedItems = List<String>();
      for (int i = 0; i < _spanMap[sheet].length; i++) {
        _Span spanObj = _spanMap[sheet][i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        if (columnIndex <= endColumn) {
          _Span newSpanObj = _Span();
          if (columnIndex <= startColumn) {
            startColumn += 1;
          }
          endColumn += 1;
          newSpanObj._start = [startRow, startColumn];
          newSpanObj._end = [endRow, endColumn];
          _spanMap[sheet][i] = newSpanObj;
          updateSpanCell = true;
          _mergeChanges = true;
        }
        String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!spannedItems.contains(rc)) {
          spannedItems.add(rc);
        }
      }
      _spannedItems[sheet] = spannedItems;
    }

    if (updateSpanCell) {
      _mergeChangeLookup = sheet;
    }

    if (_spanMap.containsKey(sheet) && _cellStyleOther.containsKey(sheet)) {
      Map<String, CellStyle> newColorMap = Map<String, CellStyle>();
      _cellStyleOther[sheet].forEach((key, value) {
        List l = cellCoordsFromCellId(key);
        int startRow = l[0], startColumn = l[1];
        String newKey = key;
        if (startColumn >= columnIndex) {
          newKey = getCellId(startColumn + 1, startRow);
        }
        newColorMap[newKey] = value;
      });
      _cellStyleOther[sheet] = Map.from(newColorMap);
    }

    var table = _tables[sheet];
    int columnLength = _tables[sheet]._maxCols;
    if (columnIndex >= columnLength) {
      table.rows.forEach((row) {
        int len = columnLength;
        while (len <= columnIndex) {
          row.insert(len, null);
          len++;
        }
      });
      table._maxCols += columnIndex - columnLength + 1;
    } else {
      table.rows.forEach((row) => row.insert(columnIndex, null));
      table._maxCols++;
    }
  }

  /// Remove column in [sheet] at position [columnIndex]
  removeColumn(String sheet, int columnIndex) {
    if (!_sheets.containsKey(sheet) || columnIndex >= _tables[sheet]._maxCols) {
      return;
    }
    if (!_update) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }
    if (columnIndex < 0) {
      throw RangeError.range(columnIndex, 0, _tables[sheet]._maxCols - 1);
    }

    List<String> deleteSingleColColor = List<String>();

    bool updateSpanCell = false;

    if (_spanMap.containsKey(sheet)) {
      List spannedItems = List<String>();
      for (int i = 0; i < _spanMap[sheet].length; i++) {
        _Span spanObj = _spanMap[sheet][i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        String clr = getCellId(startColumn, startRow);

        if (columnIndex <= endColumn) {
          _Span newSpanObj = _Span();
          if (columnIndex < startColumn) {
            startColumn -= 1;
          }
          endColumn -= 1;
          if (startColumn >= endColumn) {
            if (!deleteSingleColColor.contains(clr))
              deleteSingleColColor.add(clr);
            _spanMap[sheet][i] = null;
          } else {
            newSpanObj._start = [startRow, startColumn];
            newSpanObj._end = [endRow, endColumn];
            _spanMap[sheet][i] = newSpanObj;
          }
          updateSpanCell = true;
          _mergeChanges = true;
        }

        if (_spanMap[sheet][i] != null) {
          String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
          if (!spannedItems.contains(rc)) {
            spannedItems.add(rc);
          }
        }
      }
      _spannedItems[sheet] = spannedItems;
    }
    _cleanUpSpanMap(sheet);

    if (updateSpanCell) {
      _mergeChangeLookup = sheet;
    }

    if (_spanMap.containsKey(sheet) && _cellStyleOther.containsKey(sheet)) {
      Map<String, CellStyle> newColorMap = Map<String, CellStyle>();
      _cellStyleOther[sheet].forEach((key, value) {
        List l = cellCoordsFromCellId(key);
        int startRow = l[0], startColumn = l[1];
        String newKey = key;
        if (!deleteSingleColColor.contains(key)) {
          if (startColumn > columnIndex) {
            newKey = getCellId(startColumn - 1, startRow);
          }
          newColorMap[newKey] = value;
        }
      });
      _cellStyleOther[sheet] = Map.from(newColorMap);
    }

    var table = _tables[sheet];
    table.rows.forEach((row) => row.removeAt(columnIndex));
    table._maxCols--;
  }

  /// Insert row in [sheet] at position [rowIndex]
  insertRow(String sheet, int rowIndex) {
    _checkSheetArguments(sheet);
    _checkSheetMaxRow(sheet, rowIndex);

    if (rowIndex < 0) {
      throw RangeError.range(rowIndex, 0, _tables[sheet]._maxRows);
    }

    bool updateSpanCell = false;

    if (_spanMap.containsKey(sheet)) {
      List spannedItems = List<String>();
      for (int i = 0; i < _spanMap[sheet].length; i++) {
        _Span spanObj = _spanMap[sheet][i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        if (rowIndex <= endRow) {
          _Span newSpanObj = _Span();
          if (rowIndex <= startRow) {
            startRow += 1;
          }
          endRow += 1;
          newSpanObj._start = [startRow, startColumn];
          newSpanObj._end = [endRow, endColumn];
          _spanMap[sheet][i] = newSpanObj;
          updateSpanCell = true;
          _mergeChanges = true;
        }
        String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
        if (!spannedItems.contains(rc)) {
          spannedItems.add(rc);
        }
      }
      _spannedItems[sheet] = spannedItems;
    }

    if (updateSpanCell) {
      _mergeChangeLookup = sheet;
    }

    if (_spanMap.containsKey(sheet) && _cellStyleOther.containsKey(sheet)) {
      Map<String, CellStyle> newColorMap = Map<String, CellStyle>();
      _cellStyleOther[sheet].forEach((key, value) {
        List l = cellCoordsFromCellId(key);
        int startRow = l[0], startColumn = l[1];
        String newKey = key;
        if (startRow >= rowIndex) {
          newKey = getCellId(startColumn, startRow + 1);
        }
        newColorMap[newKey] = value;
      });
      _cellStyleOther[sheet] = Map.from(newColorMap);
    }

    var table = _tables[sheet];
    if (rowIndex >= _tables[sheet]._maxRows) {
      while (_tables[sheet]._maxRows <= rowIndex) {
        table.rows.insert(_tables[sheet]._maxRows,
            List.generate(table._maxCols, (_) => null));
        table._maxRows++;
      }
    } else {
      table.rows.insert(rowIndex, List.generate(table._maxCols, (_) => null));
      table._maxRows++;
    }
  }

  /// Remove row in [sheet] at position [rowIndex]
  removeRow(String sheet, int rowIndex) {
    if (!_sheets.containsKey(sheet) || rowIndex >= _tables[sheet]._maxRows) {
      return;
    }
    if (!_update) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }
    if (rowIndex < 0) {
      throw RangeError.range(rowIndex, 0, _tables[sheet]._maxRows - 1);
    }
    List<String> deleteSingleRowColor = List<String>();

    bool updateSpanCell = false;

    if (_spanMap != null && _spanMap.containsKey(sheet)) {
      List spannedItems = List<String>();
      for (int i = 0; i < _spanMap[sheet].length; i++) {
        _Span spanObj = _spanMap[sheet][i];
        int startColumn = spanObj.columnSpanStart,
            startRow = spanObj.rowSpanStart,
            endColumn = spanObj.columnSpanEnd,
            endRow = spanObj.rowSpanEnd;

        String clr = getCellId(startColumn, startRow);

        if (rowIndex <= endRow) {
          _Span newSpanObj = _Span();
          if (rowIndex < startRow) {
            startRow -= 1;
          }
          endRow -= 1;
          if (startRow >= endRow) {
            if (!deleteSingleRowColor.contains(clr))
              deleteSingleRowColor.add(clr);
            _spanMap[sheet][i] = null;
          } else {
            newSpanObj._start = [startRow, startColumn];
            newSpanObj._end = [endRow, endColumn];
            _spanMap[sheet][i] = newSpanObj;
          }
          updateSpanCell = true;
          _mergeChanges = true;
        }
        if (_spanMap[sheet][i] != null) {
          String rc = _getSpanCellId(startColumn, startRow, endColumn, endRow);
          if (!spannedItems.contains(rc)) {
            spannedItems.add(rc);
          }
        }
      }
      _spannedItems[sheet] = spannedItems;
    }

    _cleanUpSpanMap(sheet);

    if (updateSpanCell) {
      _mergeChangeLookup = sheet;
    }

    if (_spanMap.containsKey(sheet) && _cellStyleOther.containsKey(sheet)) {
      Map<String, CellStyle> newColorMap = Map<String, CellStyle>();
      _cellStyleOther[sheet].forEach((key, value) {
        List l = cellCoordsFromCellId(key);
        int startRow = l[0], startColumn = l[1];
        String newKey = key;
        if (!deleteSingleRowColor.contains(key)) {
          if (startRow > rowIndex) {
            newKey = getCellId(startColumn, startRow - 1);
          }
          newColorMap[newKey] = value;
        }
      });
      _cellStyleOther[sheet] = Map.from(newColorMap);
    }

    var table = _tables[sheet];
    table.rows.removeAt(rowIndex);
    table._maxRows--;
  }

  List<int> _isInsideSpanning(String sheet, int rowIndex, int columnIndex) {
    int newRowIndex = rowIndex, newColumnIndex = columnIndex;

    if (_spanMap.containsKey(sheet) && _spanMap[sheet].isNotEmpty) {
      for (int i = 0; i < _spanMap[sheet].length; i++) {
        _Span spanObj = _spanMap[sheet][i];

        if (rowIndex >= spanObj.rowSpanStart &&
            rowIndex <= spanObj.rowSpanEnd &&
            columnIndex >= spanObj.columnSpanStart &&
            columnIndex <= spanObj.columnSpanEnd) {
          newRowIndex = spanObj.rowSpanStart;
          newColumnIndex = spanObj.columnSpanStart;
          break;
        }
      }
    }

    return [newRowIndex, newColumnIndex];
  }

  /// Update the contents from sheet of the cell index: [columnIndex , rowIndex] where indexing starts from 0
  ///
  /// --or-- by Cell-Id: "A1"
  ///
  /// Font / Background color can be updated by providing Hex String to [fontColorHex] / [backgroundColorHex] as required.
  updateCell(String sheet, CellIndex cellIndex, dynamic value,
      {String fontColorHex,
      String backgroundColorHex,
      TextWrapping wrap,
      VerticalAlign verticalAlign,
      HorizontalAlign horizontalAlign}) {
    _checkSheetArguments(sheet);
    int columnIndex = cellIndex._columnIndex;
    int rowIndex = cellIndex._rowIndex;

    if (columnIndex >= _tables[sheet]._maxCols) {
      insertColumn(sheet, columnIndex);
    }

    if (rowIndex >= _tables[sheet]._maxRows) {
      insertRow(sheet, rowIndex);
    }

    int newRowIndex = rowIndex, newColumnIndex = columnIndex;
    if (_spanMap.containsKey(sheet)) {
      List updatedPosition = _isInsideSpanning(sheet, rowIndex, columnIndex);
      newRowIndex = updatedPosition[0];
      newColumnIndex = updatedPosition[1];
    }

    _tables[sheet].rows[newRowIndex][newColumnIndex] = value.toString();
    if (!_sharedStrings.contains('$value')) {
      _sharedStrings.add(value.toString());
    }
  }

  void merge(String sheet, CellIndex start, CellIndex end,
      {dynamic customValue}) {
    _checkSheetArguments(sheet);
    int startColumn = start._columnIndex,
        startRow = start._rowIndex,
        endColumn = end._columnIndex,
        endRow = end._rowIndex;

    _checkSheetMaxCol(sheet, startColumn);
    _checkSheetMaxCol(sheet, endColumn);
    _checkSheetMaxRow(sheet, startRow);
    _checkSheetMaxRow(sheet, endRow);

    if ((startColumn == endColumn && startRow == endRow) ||
        (startColumn < 0 || startRow < 0 || endColumn < 0 || endRow < 0) ||
        (_spannedItems != null &&
            _spannedItems.containsKey(sheet) &&
            _spannedItems[sheet].contains(
                _getSpanCellId(startColumn, startRow, endColumn, endRow)))) {
      return;
    }

    List<int> gotPosition = _getSpanPosition(sheet, start, end);

    String value;
    bool gotValue = true;

    if (customValue != null) {
      value = '$customValue';
      gotValue = false;
    }
    _mergeChanges = true;

    startColumn = gotPosition[0];
    startRow = gotPosition[1];
    endColumn = gotPosition[2];
    endRow = gotPosition[3];

    for (int j = startRow; j <= endRow; j++) {
      for (int k = startColumn; k <= endColumn; k++) {
        if ((_tables != null && !_tables.containsKey(sheet)) ||
            j == _tables[sheet].maxRows ||
            k == _tables[sheet].maxCols) {
          updateCell(sheet,
              CellIndex.indexByColumnRow(columnIndex: k, rowIndex: j), null);
        } else {
          if (gotValue && _tables[sheet].rows[j][k] != null) {
            value = "${_tables[sheet].rows[j][k]}";
            gotValue = false;
          }
          _tables[sheet].rows[j][k] = null;
        }
      }
    }

    _tables[sheet].rows[startRow][startColumn] = value;

    String sp = _getSpanCellId(startColumn, startRow, endColumn, endRow);
    List<String> ls = List<String>();

    if (_spannedItems != null && _spannedItems.containsKey(sheet)) {
      ls = List<String>.of(_spannedItems[sheet]);
    }
    if (!ls.contains(sp)) {
      ls.add(sp);
    }

    _spannedItems[sheet] = ls;

    _tables[sheet].rows[startRow][startColumn] = value;
    List l = List<_Span>();
    _Span s = _Span();
    s._start = [startRow, startColumn];
    s._end = [endRow, endColumn];

    if (_spanMap.containsKey(sheet) && _spanMap[sheet].isNotEmpty) {
      l = List<_Span>.of(_spanMap[sheet]);
    }
    l.add(s);
    _spanMap[sheet] = l;
    _mergeChangeLookup = sheet;
  }

  Map<String, List<int>> _isLocationChangeRequired(String sheet,
      int startColumn, int startRow, int endColumn, int endRow, _Span spanObj) {
    int changeValue = 0;

    if (startRow <= spanObj.rowSpanStart &&
        startColumn <= spanObj.columnSpanStart &&
        endRow >= spanObj.rowSpanEnd &&
        endColumn >= spanObj.columnSpanEnd) {
      changeValue = 1;
    } else {
      if ((startColumn < spanObj.columnSpanStart &&
              endColumn >= spanObj.columnSpanStart) ||
          (startColumn <= spanObj.columnSpanEnd &&
              endColumn > spanObj.columnSpanEnd)) {
        /**
           * Start Row stretching to up position
           */
        if (startRow >= spanObj.rowSpanStart &&
            startRow <= spanObj.rowSpanEnd) {
          startRow = spanObj.rowSpanStart;
          changeValue = 1;
        }
        /**
           * End Row stretching to bottom position
           */
        if (endRow >= spanObj.rowSpanStart && endRow <= spanObj.rowSpanEnd) {
          endRow = spanObj.rowSpanEnd;
          changeValue = 1;
        }

        if (startColumn >= spanObj.columnSpanStart) {
          startColumn = spanObj.columnSpanStart;
          changeValue = 1;
        }

        if (endColumn <= spanObj.columnSpanEnd) {
          endColumn = spanObj.columnSpanEnd;
          changeValue = 1;
        }
      }

      if ((startRow < spanObj.rowSpanStart && endRow >= spanObj.rowSpanStart) ||
          (startRow <= spanObj.rowSpanEnd && endRow > spanObj.rowSpanEnd)) {
        /**
           * Start Column stretching to left positionc
           */
        if (startColumn >= spanObj.columnSpanStart &&
            startColumn <= spanObj.columnSpanEnd) {
          startColumn = spanObj.columnSpanStart;
          changeValue = 1;
        }
        /**
           * End Column stretching to right position
           */
        if (endColumn >= spanObj.columnSpanStart &&
            endColumn <= spanObj.columnSpanEnd) {
          endColumn = spanObj.columnSpanEnd;
          changeValue = 1;
        }

        if (startRow >= spanObj.rowSpanStart) {
          startRow = spanObj.rowSpanStart;
          changeValue = 1;
        }

        if (endRow <= spanObj.rowSpanEnd) {
          endRow = spanObj.rowSpanEnd;
          changeValue = 1;
        }
      }
    }

    return Map<String, List<int>>.from({
      "gotPosition": List<int>.from([changeValue]),
      "changeValue": List<int>.from([startColumn, startRow, endColumn, endRow])
    });
  }

  /// Helps to find the interaction between the pre-existing span position
  /// and updates if with new span if there any interaction(Cross-Sectional Spanning) exists.
  List<int> _getSpanPosition(String sheet, CellIndex start, CellIndex end) {
    int startColumn = start._columnIndex,
        startRow = start._rowIndex,
        endColumn = end._columnIndex,
        endRow = end._rowIndex;

    bool remove = false;

    if (startRow > endRow) {
      startRow = end._rowIndex;
      endRow = start._rowIndex;
    }
    if (endColumn < startColumn) {
      endColumn = start._columnIndex;
      startColumn = end._columnIndex;
    }

    if (_spanMap != null &&
        _spanMap.containsKey(sheet) &&
        _spanMap[sheet].isNotEmpty) {
      List<_Span> data = _spanMap[sheet];
      for (int i = 0; i < data.length; i++) {
        _Span spanObj = data[i];

        Map<String, List<int>> gotMap = _isLocationChangeRequired(
            sheet, startColumn, startRow, endColumn, endRow, spanObj);
        List<int> gotPosition = gotMap['gotPosition'],
            changeValue = gotMap['changeValue'];

        if (changeValue == 1) {
          startColumn = gotPosition[0];
          startRow = gotPosition[1];
          endColumn = gotPosition[2];
          endRow = gotPosition[3];
          String sp = _getSpanCellId(spanObj.columnSpanStart,
              spanObj.rowSpanStart, spanObj.columnSpanEnd, spanObj.rowSpanEnd);
          if (_spannedItems != null &&
              _spannedItems.containsKey(sheet) &&
              _spannedItems[sheet].contains(sp)) {
            _spannedItems[sheet].remove(sp);
          }
          remove = true;
          _spanMap[sheet][i] = null;
        }
      }
      if (remove) {
        _cleanUpSpanMap(sheet);
      }
    }
    return [startColumn, startRow, endColumn, endRow];
  }

  String getCellId(int colI, int rowI) =>
      '${numericToLetters(colI + 1)}${rowI + 1}';

  String _getSpanCellId(
      int startColumn, int startRow, int endColumn, int endRow) {
    return '${getCellId(startColumn, startRow)}:${getCellId(endColumn, endRow)}';
  }

  /// returns an Iterable of cell-Id for the previously merged cell-Ids.
  Iterable<String> getMergedCells(String sheet) {
    return _spannedItems != null && _spannedItems.containsKey(sheet)
        ? List<String>.of(_spannedItems[sheet])
        : [];
  }

  /**
   * Usage this function to unMerge the merged cells.
   * 
   *        var sheet = 'DesiredSheet';
   *        List<String> spannedCells = excel.getMergedCells(sheet);
   *        var cellToUnMerge = "A1:A2";
   *        excel.unMerge(sheet, cellToUnMerge);
   * 
   */
  unMerge(String sheet, String unmergeCells) {
    if (_spannedItems != null &&
        _spannedItems.containsKey(sheet) &&
        _spanMap != null &&
        _spanMap.containsKey(sheet) &&
        unmergeCells != null &&
        _spannedItems[sheet].contains(unmergeCells)) {
      List<String> lis = unmergeCells.split(RegExp(r":"));
      if (lis.length == 2) {
        bool remove = false;
        List<int> start, end;
        start =
            cellCoordsFromCellId(lis[0]); // [x,y] => [startRow, startColumn]
        end = cellCoordsFromCellId(lis[1]); // [x,y] => [endRow, endColumn]
        for (int i = 0; i < _spanMap[sheet].length; i++) {
          _Span spanObject = _spanMap[sheet][i];

          if (spanObject.columnSpanStart == start[1] &&
              spanObject.rowSpanStart == start[0] &&
              spanObject.columnSpanEnd == end[1] &&
              spanObject.rowSpanEnd == end[0]) {
            _spanMap[sheet][i] = null;
            remove = true;
          }
        }
        if (remove) {
          _cleanUpSpanMap(sheet);
        }
      }
      _spannedItems[sheet].remove(unmergeCells);
      _mergeChangeLookup = sheet;
    }
  }

  // Cleaning up the null values from the Span Map
  _cleanUpSpanMap(String sheet) {
    _spanMap[sheet].removeWhere((value) => value == null);
    if (_spanMap.containsKey(sheet) && _spanMap[sheet].length < 1) {
      _spanMap.remove(sheet);
    }
  }

  bool _isEmptyRow(List row) {
    return row.fold(true, (value, element) => value && (element == null));
  }

  bool _isNotEmptyRow(List row) {
    return !_isEmptyRow(row);
  }

  set _mergeChangeLookup(String value) {
    if (!_mergeChangeLook.contains(value)) {
      _mergeChangeLook.add(value);
    }
  }

  _countFilledRow(DataTable table, List row) {
    if (_isNotEmptyRow(row) && table._maxRows < table._rows.length) {
      table._maxRows = table._rows.length;
    }
  }

  _countFilledColumn(DataTable table, List row, dynamic value) {
    if (value != null && table._maxCols < row.length) {
      table._maxCols = row.length;
    }
  }

  _parseTable(XmlElement node) {
    var name = node.getAttribute('name');
    var target = _worksheetTargets[node.getAttribute('r:id')];

    tables[name] = DataTable(name);
    var table = tables[name];

    var file = _archive.findFile('xl/$target');
    file.decompress();

    var content = parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, table, name);
    });

    if (_update) {
      _sheets[name] = sheet;
      _xmlFiles['xl/$target'] = content;
    }
    _xmlSheetId[name] = 'xl/$target';

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, DataTable table, String name) {
    var row = List();

    _findCells(node).forEach((child) {
      _parseCell(child, table, row, name);
    });

    var rowIndex = _getRowNumber(node) - 1;
    if (_isNotEmptyRow(row) && rowIndex > table._rows.length) {
      var repeat = rowIndex - table._rows.length;
      for (var index = 0; index < repeat; index++) {
        table._rows.add(List());
      }
    }

    if (_isNotEmptyRow(row)) {
      table._rows.add(row);
    } else {
      table._rows.add(List());
    }

    _countFilledRow(table, row);
  }

  _parseCell(XmlElement node, DataTable table, List row, String name) {
    var colIndex = _getCellNumber(node);
    if (colIndex > row.length) {
      var repeat = colIndex - row.length;
      for (var index = 0; index < repeat; index++) {
        row.add(null);
      }
    }

    var s1 = node.getAttribute('s');
    int s = 0;
    if (s1 != null) {
      try {
        s = int.parse(s1.toString());
      } catch (_) {}

      String rC = node.getAttribute('r').toString();

      if (_cellStyleReferenced.containsKey(name)) {
        _cellStyleReferenced[name][rC] = s;
      } else {
        _cellStyleReferenced[name] = {rC: s};
      }
    }

    if (node.children.isEmpty) {
      return;
    }

    var value, type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        value = _sharedStrings[
            int.parse(_parseValue(node.findElements('v').first))];
        break;
      // boolean
      case 'b':
        value = _parseValue(node.findElements('v').first) == '1';
        break;
      // error
      case 'e':
      // formula
      case 'str':
        value = _parseValue(node.findElements('v').first);
        break;
      // inline string
      case 'inlineStr':
        // <c r='B2' t='inlineStr'>
        // <is><t>Dartonico</t></is>
        // </c>
        value = _parseValue(node.findAllElements('t').first);
        break;
      // number
      case 'n':
      default:
        var valueNode = node.findElements('v');
        var content = valueNode.first;
        if (s1 != null) {
          var fmtId = _numFormats[s];
          // date
          if (((fmtId >= 14) && (fmtId <= 17)) || (fmtId == 22)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(1899, 12, 30);
            value = date
                .add(Duration(milliseconds: delta.toInt()))
                .toIso8601String();
            // time
          } else if (((fmtId >= 18) && (fmtId <= 21)) ||
              ((fmtId >= 45) && (fmtId <= 47))) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(0);
            date = date.add(Duration(milliseconds: delta.toInt()));
            value =
                '${_twoDigits(date.hour)}:${_twoDigits(date.minute)}:${_twoDigits(date.second)}';
            // number
          } else {
            value = num.parse(_parseValue(content));
          }
        } else {
          value = num.parse(_parseValue(content));
        }
    }
    row.add(value);
    if (!_sharedStrings.contains('$value')) {
      _sharedStrings.add('$value');
    }

    _countFilledColumn(table, row, value);
  }

  _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
      }
    });

    return buffer.toString();
  }

  Iterable<XmlElement> _findRows(XmlElement table) {
    return table.findElements('row');
  }

  Iterable<XmlElement> _findCells(XmlElement row) {
    return row.findElements('c');
  }

  int _getCellNumber(XmlElement cell) {
    return cellCoordsFromCellId(cell.getAttribute('r'))[1];
  }

  int _getRowNumber(XmlElement row) {
    return int.parse(row.getAttribute('r'));
  }

  XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = 0;
    for (var currentRow in rows) {
      currentIndex = _getRowNumber(currentRow) - 1;
      if (currentIndex >= rowIndex) {
        row = currentRow;
        break;
      }
    }

    // Create row if required
    if (row == null || currentIndex != rowIndex) {
      row = __insertRow(table, row, rowIndex);
    }

    return row;
  }

  XmlElement _updateCell(String sheet, XmlElement node, int columnIndex,
      int rowIndex, dynamic value) {
    XmlElement cell;
    var cells = _findCells(node);

    var currentIndex = 0; // cells could be empty
    for (var currentCell in cells) {
      currentIndex = _getCellNumber(currentCell);
      if (currentIndex >= columnIndex) {
        cell = currentCell;
        break;
      }
    }

    if (cell == null || currentIndex != columnIndex) {
      cell = _insertCell(sheet, node, cell, columnIndex, rowIndex, value);
    } else {
      cell = _replaceCell(sheet, node, cell, columnIndex, rowIndex, value);
    }

    return cell;
  }

  XmlElement _createRow(int rowIndex) => XmlElement(XmlName('row'),
      [XmlAttribute(XmlName('r'), (rowIndex + 1).toString())], []);

  XmlElement __insertRow(XmlElement table, XmlElement lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null) {
      table.children.add(row);
    } else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }

  XmlElement _insertCell(String sheet, XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    if (lastCell == null) {
      row.children.add(cell);
    } else {
      var index = row.children.indexOf(lastCell);
      row.children.insert(index, cell);
    }
    return cell;
  }

  XmlElement _replaceCell(String sheet, XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var index = lastCell == null ? 0 : row.children.indexOf(lastCell);
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  }

  // Manage value's type
  XmlElement _createCell(
      String sheet, int columnIndex, int rowIndex, dynamic value) {
    if (!_sharedStrings.contains(value.toString())) {
      _sharedStrings.add(value.toString());
    }

    String rC = getCellId(columnIndex, rowIndex);

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      XmlAttribute(XmlName('t'), 's'),
    ];

    if (_colorChanges &&
        _cellStyleOther.containsKey(sheet) &&
        _cellStyleOther[sheet].containsKey(rC)) {
      CellStyle cellStyle = _cellStyleOther[sheet][rC];
      int upperLevelPos = _checkPosition(_cellStyleList, cellStyle);
      if (upperLevelPos == -1) {
        int lowerLevelPos = _checkPosition(_innerCellStyle, cellStyle);
        if (lowerLevelPos != -1) {
          upperLevelPos = lowerLevelPos + _cellStyleList.length;
        } else {
          upperLevelPos = 0;
        }
      }
      attributes.insert(
        1,
        XmlAttribute(XmlName('s'), '$upperLevelPos'),
      );
    } else if (_colorChanges &&
        _cellStyleReferenced.containsKey(sheet) &&
        _cellStyleReferenced[sheet].containsKey(rC)) {
      attributes.insert(
        1,
        XmlAttribute(XmlName('s'), '${_cellStyleReferenced[sheet][rC]}'),
      );
    }
    var children = value == null
        ? <XmlElement>[]
        : <XmlElement>[
            XmlElement(XmlName('v'), [],
                [XmlText(_sharedStrings.indexOf(value.toString()).toString())]),
          ];
    return XmlElement(XmlName('c'), attributes, children);
  }
}
