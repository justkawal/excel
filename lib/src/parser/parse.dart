part of excel;

class Parser {
  final Excel _excel;
  final List<String> _rId = [];
  final Map<String, String> _worksheetTargets = {};

  Parser._(this._excel);

  void _startParsing() {
    _putContentXml();
    _parseRelations();
    _parseStyles(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _parseMergedCells();
  }

  void _normalizeTable(Sheet sheet) {
    if (sheet._maxRows == 0 || sheet._maxColumns == 0) {
      sheet._sheetData.clear();
    }
    sheet._countRowsAndColumns();
  }

  void _putContentXml() {
    var file = _excel._archive.findFile("[Content_Types].xml");

    if (file == null) {
      _damagedExcel();
    }
    file!.decompress();
    _excel._xmlFiles["[Content_Types].xml"] =
        XmlDocument.parse(utf8.decode(file.content));
  }

  void _parseRelations() {
    var relations = _excel._archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = XmlDocument.parse(utf8.decode(relations.content));
      _excel._xmlFiles['xl/_rels/workbook.xml.rels'] = document;

      document.findAllElements('Relationship').forEach((node) {
        String? id = node.getAttribute('Id');
        String? target = node.getAttribute('Target');
        if (target != null) {
          switch (node.getAttribute('Type')) {
            case _relationshipsStyles:
              _excel._stylesTarget = target;
              break;
            case _relationshipsWorksheet:
              if (id != null) _worksheetTargets[id] = target;
              break;
            case _relationshipsSharedStrings:
              _excel._sharedStringsTarget = target;
              break;
          }
        }
        if (id != null && !_rId.contains(id)) {
          _rId.add(id);
        }
      });
    } else {
      _damagedExcel();
    }
  }

  void _parseSharedStrings() {
    var sharedStrings =
        _excel._archive.findFile(_excel._absSharedStringsTarget);
    if (sharedStrings == null) {
      _excel._sharedStringsTarget = 'sharedStrings.xml';

      /// Running it with false will collect all the `rid` and will
      /// help us to get the available rid to assign it to `sharedStrings.xml` back
      _parseContent(run: false);

      if (_excel._xmlFiles.containsKey("xl/_rels/workbook.xml.rels")) {
        int rIdNumber = _getAvailableRid();

        _excel._xmlFiles["xl/_rels/workbook.xml.rels"]
            ?.findAllElements('Relationships')
            .first
            .children
            .add(XmlElement(
              XmlName('Relationship'),
              <XmlAttribute>[
                XmlAttribute(XmlName('Id'), 'rId$rIdNumber'),
                XmlAttribute(XmlName('Type'),
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'),
                XmlAttribute(XmlName('Target'), 'sharedStrings.xml')
              ],
            ));
        if (!_rId.contains('rId$rIdNumber')) {
          _rId.add('rId$rIdNumber');
        }
        String content =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        bool contain = true;

        _excel._xmlFiles["[Content_Types].xml"]
            ?.findAllElements('Override')
            .forEach((node) {
          var value = node.getAttribute('ContentType');
          if (value == content) {
            contain = false;
          }
        });
        if (contain) {
          _excel._xmlFiles["[Content_Types].xml"]
              ?.findAllElements('Types')
              .first
              .children
              .add(XmlElement(
                XmlName('Override'),
                <XmlAttribute>[
                  XmlAttribute(XmlName('PartName'), '/xl/sharedStrings.xml'),
                  XmlAttribute(XmlName('ContentType'), content),
                ],
              ));
        }
      }

      var content = utf8.encode(
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>");
      _excel._archive.addFile(
          ArchiveFile("xl/sharedStrings.xml", content.length, content));
      sharedStrings = _excel._archive.findFile("xl/sharedStrings.xml");
    }
    sharedStrings!.decompress();
    var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
    _excel._xmlFiles["xl/${_excel._sharedStringsTarget}"] = document;

    document.findAllElements('si').forEach((node) {
      _parseSharedString(node);
    });
  }

  void _parseSharedString(XmlElement node) {
    final sharedString = SharedString(node: node);
    _excel._sharedStrings.add(sharedString, sharedString.stringValue);
  }

  void _parseContent({bool run = true}) {
    var workbook = _excel._archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      _damagedExcel();
    }
    workbook!.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook.content));
    _excel._xmlFiles["xl/workbook.xml"] = document;

    document.findAllElements('sheet').forEach((node) {
      if (run) {
        _parseTable(node);
      } else {
        var rid = node.getAttribute('r:id');
        if (rid != null && !_rId.contains(rid)) {
          _rId.add(rid);
        }
      }
    });
  }

  /// Parses and processes merged cells within the spreadsheet.
  ///
  /// This method identifies merged cell regions in each sheet of the spreadsheet
  /// and handles them accordingly. It removes all cells within a merged cell region
  /// except for the top-left cell, preserving its content.
  void _parseMergedCells() {
    Map spannedCells = <String, List<String>>{};
    _excel._sheets.forEach((sheetName, node) {
      _excel._availSheet(sheetName);
      XmlElement sheetDataNode = node as XmlElement;
      List spanList = <String>[];
      final sheet = _excel._sheetMap[sheetName]!;

      final worksheetNode = sheetDataNode.parent;
      worksheetNode!.findAllElements('mergeCell').forEach((element) {
        String? ref = element.getAttribute('ref');
        if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
          if (!sheet._spannedItems.contains(ref)) {
            sheet._spannedItems.add(ref);
          }

          String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

          if (!spanList.contains(startCell)) {
            spanList.add(startCell);
          }
          spannedCells[sheetName] = spanList;

          CellIndex startIndex = CellIndex.indexByString(startCell),
              endIndex = CellIndex.indexByString(endCell);
          _Span spanObj = _Span.fromCellIndex(
            start: startIndex,
            end: endIndex,
          );
          if (!sheet._spanList.contains(spanObj)) {
            sheet._spanList.add(spanObj);

            _deleteAllButTopLeftCellsOfSpanObj(spanObj, sheet);
          }
          _excel._mergeChangeLookup = sheetName;
        }
      });
    });
  }

  /// Deletes all cells within the span of the given [_Span] object
  /// except for the top-left cell.
  ///
  /// This method is used internally by [_parseMergedCells] to remove
  /// cells within merged cell regions.
  ///
  /// Parameters:
  ///   - [spanObj]: The span object representing the merged cell region.
  ///   - [sheet]: The sheet object from which cells are to be removed.
  void _deleteAllButTopLeftCellsOfSpanObj(_Span spanObj, Sheet sheet) {
    final columnSpanStart = spanObj.columnSpanStart;
    final columnSpanEnd = spanObj.columnSpanEnd;
    final rowSpanStart = spanObj.rowSpanStart;
    final rowSpanEnd = spanObj.rowSpanEnd;

    for (var columnI = columnSpanStart; columnI <= columnSpanEnd; columnI++) {
      for (var rowI = rowSpanStart; rowI <= rowSpanEnd; rowI++) {
        bool isTopLeftCellThatShouldNotBeDeleted =
            columnI == columnSpanStart && rowI == rowSpanStart;

        if (isTopLeftCellThatShouldNotBeDeleted) {
          continue;
        }
        sheet._removeCell(rowI, columnI);
      }
    }
  }

  // Reading the styles from the excel file.
  void _parseStyles(String _stylesTarget) {
    var styles = _excel._archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));
      _excel._xmlFiles['xl/$_stylesTarget'] = document;

      _excel._fontStyleList = <_FontStyle>[];
      _excel._patternFill = <String>[];
      _excel._cellStyleList = <CellStyle>[];
      _excel._borderSetList = <_BorderSet>[];

      Iterable<XmlElement> fontList = document.findAllElements('font');

      document.findAllElements('patternFill').forEach((node) {
        String patternType = node.getAttribute('patternType') ?? '', rgb;
        if (node.children.isNotEmpty) {
          node.findElements('fgColor').forEach((child) {
            rgb = child.getAttribute('rgb') ?? '';
            _excel._patternFill.add(rgb);
          });
        } else {
          _excel._patternFill.add(patternType);
        }
      });

      document.findAllElements('border').forEach((node) {
        final diagonalUp = !['0', 'false', null]
            .contains(node.getAttribute('diagonalUp')?.trim());
        final diagonalDown = !['0', 'false', null]
            .contains(node.getAttribute('diagonalDown')?.trim());

        const List<String> borderElementNamesList = [
          'left',
          'right',
          'top',
          'bottom',
          'diagonal'
        ];
        Map<String, Border> borderElements = {};
        for (var elementName in borderElementNamesList) {
          XmlElement? element;
          try {
            element = node.findElements(elementName).single;
          } on StateError catch (_) {
            // Either there is no element, or there are too many ones.
            // Silently ignore this element.
          }

          final borderStyleAttribute = element?.getAttribute('style')?.trim();
          final borderStyle = borderStyleAttribute != null
              ? getBorderStyleByName(borderStyleAttribute)
              : null;

          String? borderColorHex;
          try {
            final color = element?.findElements('color').single;
            borderColorHex = color?.getAttribute('rgb')?.trim();
          } on StateError catch (_) {}

          borderElements[elementName] = Border(
              borderStyle: borderStyle,
              borderColorHex: borderColorHex?.excelColor);
        }

        final borderSet = _BorderSet(
          leftBorder: borderElements['left']!,
          rightBorder: borderElements['right']!,
          topBorder: borderElements['top']!,
          bottomBorder: borderElements['bottom']!,
          diagonalBorder: borderElements['diagonal']!,
          diagonalBorderDown: diagonalDown,
          diagonalBorderUp: diagonalUp,
        );
        _excel._borderSetList.add(borderSet);
      });

      document.findAllElements('numFmts').forEach((node1) {
        node1.findAllElements('numFmt').forEach((node) {
          var numFmtId = int.parse(node.getAttribute('numFmtId')!);
          final formatCode = node.getAttribute('formatCode')!;
          if (numFmtId < 164) {
            numFmtId += 164;
            // throw Exception(
            //     'custom numFmtId starts at 164 but found a value of $numFmtId');
          }

          _excel._numFormats
              .add(numFmtId, NumFormat.custom(formatCode: formatCode));
        });
      });

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          final numFmtId = _getFontIndex(node, 'numFmtId');
          _excel._numFmtIds.add(numFmtId);

          String fontColor = ExcelColor.black.colorHex,
              backgroundColor = ExcelColor.none.colorHex;
          String? fontFamily;
          FontScheme fontScheme = FontScheme.Unset;
          _BorderSet? borderSet;

          int fontSize = 12;
          bool isBold = false, isItalic = false;
          Underline underline = Underline.None;
          HorizontalAlign horizontalAlign = HorizontalAlign.Left;
          VerticalAlign verticalAlign = VerticalAlign.Bottom;
          TextWrapping? textWrapping;
          int rotation = 0;
          int fontId = _getFontIndex(node, 'fontId');
          _FontStyle _fontStyle = _FontStyle();

          /// checking for other font values
          if (fontId < fontList.length) {
            XmlElement font = fontList.elementAt(fontId);

            /// Checking for font Size.
            var _clr = _nodeChildren(font, 'color', attribute: 'rgb');
            if (_clr != null && !(_clr is bool)) {
              fontColor = _clr.toString();
            }

            /// Checking for font Size.
            String? _size = _nodeChildren(font, 'sz', attribute: 'val');
            if (_size != null) {
              fontSize = double.parse(_size).round();
            }

            /// Checking for bold
            var _bold = _nodeChildren(font, 'b');
            if (_bold != null && _bold is bool && _bold) {
              isBold = true;
            }

            /// Checking for italic
            var _italic = _nodeChildren(font, 'i');
            if (_italic != null && _italic) {
              isItalic = true;
            }

            /// Checking for double underline
            var _underline = _nodeChildren(font, 'u', attribute: 'val');
            if (_underline != null) {
              underline = Underline.Double;
            }

            /// Checking for single underline
            var _singleUnderline = _nodeChildren(font, 'u');
            if (_singleUnderline != null) {
              underline = Underline.Single;
            }

            /// Checking for font Family
            var _family = _nodeChildren(font, 'name', attribute: 'val');
            if (_family != null && _family != true) {
              fontFamily = _family;
            }

            /// Checking for font Scheme
            var _scheme = _nodeChildren(font, 'scheme', attribute: 'val');
            if (_scheme != null) {
              fontScheme =
                  _scheme == "major" ? FontScheme.Major : FontScheme.Minor;
            }

            _fontStyle.isBold = isBold;
            _fontStyle.isItalic = isItalic;
            _fontStyle.fontSize = fontSize;
            _fontStyle.fontFamily = fontFamily;
            _fontStyle.fontScheme = fontScheme;
            _fontStyle._fontColorHex = fontColor.excelColor;
          }

          /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fontStyleList`
          if (_fontStyleIndex(_excel._fontStyleList, _fontStyle) == -1) {
            _excel._fontStyleList.add(_fontStyle);
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _excel._patternFill.length) {
            backgroundColor = _excel._patternFill[fillId];
          }

          int borderId = _getFontIndex(node, 'borderId');
          if (borderId < _excel._borderSetList.length) {
            borderSet = _excel._borderSetList[borderId];
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

              var rotationString = node.getAttribute('textRotation');
              if (rotationString != null) {
                rotation = (double.tryParse(rotationString) ?? 0.0).floor();
              }
            });
          }

          var numFormat = _excel._numFormats.getByNumFmtId(numFmtId);
          if (numFormat == null) {
            assert(false, 'missing numFmt for $numFmtId');
            numFormat = NumFormat.standard_0;
          }

          CellStyle cellStyle = CellStyle(
            fontColorHex: fontColor.excelColor,
            fontFamily: fontFamily,
            fontSize: fontSize,
            bold: isBold,
            italic: isItalic,
            underline: underline,
            backgroundColorHex:
                backgroundColor == 'none' || backgroundColor.isEmpty
                    ? ExcelColor.none
                    : backgroundColor.excelColor,
            horizontalAlign: horizontalAlign,
            verticalAlign: verticalAlign,
            textWrapping: textWrapping,
            rotation: rotation,
            leftBorder: borderSet?.leftBorder,
            rightBorder: borderSet?.rightBorder,
            topBorder: borderSet?.topBorder,
            bottomBorder: borderSet?.bottomBorder,
            diagonalBorder: borderSet?.diagonalBorder,
            diagonalBorderUp: borderSet?.diagonalBorderUp ?? false,
            diagonalBorderDown: borderSet?.diagonalBorderDown ?? false,
            numberFormat: numFormat,
          );

          _excel._cellStyleList.add(cellStyle);
        });
      });
    } else {
      _damagedExcel(text: 'styles');
    }
  }

  dynamic _nodeChildren(XmlElement node, String child, {var attribute}) {
    Iterable<XmlElement> ele = node.findElements(child);
    if (ele.isNotEmpty) {
      if (attribute != null) {
        var attr = ele.first.getAttribute(attribute);
        if (attr != null) {
          return attr;
        }
        return null; // pretending that attribute is not found so sending null.
      }
      return true; // mocking to be found the children in case of bold and italic.
    }
    return null; // pretending that the node's children is not having specified child.
  }

  int _getFontIndex(var node, String text) {
    String? applyFont = node.getAttribute(text)?.trim();
    if (applyFont != null) {
      try {
        return int.parse(applyFont.toString());
      } catch (e) {
        if (applyFont.toLowerCase() == 'true') {
          return 1;
        }
      }
    }
    return 0;
  }

  void _parseTable(XmlElement node) {
    var name = node.getAttribute('name')!;
    var target = _worksheetTargets[node.getAttribute('r:id')];

    if (_excel._sheetMap['$name'] == null) {
      _excel._sheetMap['$name'] = Sheet._(_excel, '$name');
    }

    Sheet sheetObject = _excel._sheetMap['$name']!;

    var file = _excel._archive.findFile('xl/$target');
    file!.decompress();

    var content = XmlDocument.parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;

    ///
    /// check for right to left view
    ///
    var sheetView = worksheet.findAllElements('sheetView').toList();
    if (sheetView.isNotEmpty) {
      var sheetViewNode = sheetView.first;
      var rtl = sheetViewNode.getAttribute('rightToLeft');
      sheetObject.isRTL = rtl != null && rtl == '1';
    }
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, sheetObject, name);
    });

    _parseHeaderFooter(worksheet, sheetObject);
    _parseColWidthsRowHeights(worksheet, sheetObject);

    _excel._sheets[name] = sheet;

    _excel._xmlFiles['xl/$target'] = content;
    _excel._xmlSheetId[name] = 'xl/$target';

    _normalizeTable(sheetObject);
  }

  _parseRow(XmlElement node, Sheet sheetObject, String name) {
    var rowIndex = (_getRowNumber(node) ?? -1) - 1;
    if (rowIndex < 0) {
      return;
    }

    _findCells(node).forEach((child) {
      _parseCell(child, sheetObject, rowIndex, name);
    });
  }

  void _parseCell(
      XmlElement node, Sheet sheetObject, int rowIndex, String name) {
    int? columnIndex = _getCellNumber(node);
    if (columnIndex == null) {
      return;
    }

    var s1 = node.getAttribute('s');
    int s = 0;
    if (s1 != null) {
      try {
        s = int.parse(s1.toString());
      } catch (_) {}

      String rC = node.getAttribute('r').toString();

      if (_excel._cellStyleReferenced[name] == null) {
        _excel._cellStyleReferenced[name] = {rC: s};
      } else {
        _excel._cellStyleReferenced[name]![rC] = s;
      }
    }

    CellValue? value;
    String? type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        final sharedString = _excel._sharedStrings
            .value(int.parse(_parseValue(node.findElements('v').first)));
        value = TextCellValue.span(sharedString!.textSpan);
        break;
      // boolean
      case 'b':
        value = BoolCellValue(_parseValue(node.findElements('v').first) == '1');
        break;
      // error
      case 'e':
      // formula
      case 'str':
        value = FormulaCellValue(_parseValue(node.findElements('v').first));
        break;
      // inline string
      case 'inlineStr':
        // <c r='B2' t='inlineStr'>
        // <is><t>Dartonico</t></is>
        // </c>
        value = TextCellValue(_parseValue(node.findAllElements('t').first));
        break;
      // number
      case 'n':
      default:
        var formulaNode = node.findElements('f');
        if (formulaNode.isNotEmpty) {
          value = FormulaCellValue(_parseValue(formulaNode.first).toString());
        } else {
          final vNode = node.findElements('v').firstOrNull;
          if (vNode == null) {
            value = null;
          } else if (s1 != null) {
            final v = _parseValue(vNode);
            var numFmtId = _excel._numFmtIds[s];
            final numFormat = _excel._numFormats.getByNumFmtId(numFmtId);
            if (numFormat == null) {
              assert(
                  false, 'found no number format spec for numFmtId $numFmtId');
              value = NumFormat.defaultNumeric.read(v);
            } else {
              value = numFormat.read(v);
            }
          } else {
            final v = _parseValue(vNode);
            value = NumFormat.defaultNumeric.read(v);
          }
        }
    }

    sheetObject.updateCell(
      CellIndex.indexByColumnRow(columnIndex: columnIndex, rowIndex: rowIndex),
      value,
      cellStyle: _excel._cellStyleList[s],
    );
  }

  static String _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.value));
      }
    });

    return buffer.toString();
  }

  int _getAvailableRid() {
    _rId.sort((a, b) {
      return int.parse(a.substring(3)).compareTo(int.parse(b.substring(3)));
    });

    List<String> got = List<String>.from(_rId.last.split(''));
    got.removeWhere((item) {
      return !'0123456789'.split('').contains(item);
    });
    return int.parse(got.join().toString()) + 1;
  }

  ///Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  ///
  ///Creates the sheet with name `newSheet` as file output and then adds it to the archive directory.
  ///
  ///
  void _createSheet(String newSheet) {
    /*
    List<XmlNode> list = _excel._xmlFiles['xl/workbook.xml']
        .findAllElements('sheets')
        .first
        .children;
    if (list.isEmpty) {
      throw ArgumentError('');
    } */

    int _sheetId = -1;
    List<int> sheetIdList = <int>[];

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheet')
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

    for (int i = 0; i < sheetIdList.length; i++) {
      if ((i + 1) != sheetIdList[i]) {
        _sheetId = i + 1;
        break;
      }
    }
    if (_sheetId == -1) {
      if (sheetIdList.isEmpty) {
        _sheetId = 1;
      } else {
        _sheetId = sheetIdList.length + 1;
      }
    }

    int sheetNumber = _sheetId;
    int ridNumber = _getAvailableRid();

    _excel._xmlFiles['xl/_rels/workbook.xml.rels']
        ?.findAllElements('Relationships')
        .first
        .children
        .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName('Type'), '$_relationships/worksheet'),
          XmlAttribute(XmlName('Target'), 'worksheets/sheet$sheetNumber.xml'),
        ]));

    if (!_rId.contains('rId$ridNumber')) {
      _rId.add('rId$ridNumber');
    }

    _excel._xmlFiles['xl/workbook.xml']
        ?.findAllElements('sheets')
        .first
        .children
        .add(XmlElement(
          XmlName('sheet'),
          <XmlAttribute>[
            XmlAttribute(XmlName('state'), 'visible'),
            XmlAttribute(XmlName('name'), newSheet),
            XmlAttribute(XmlName('sheetId'), '$sheetNumber'),
            XmlAttribute(XmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet$sheetNumber.xml';

    var content = utf8.encode(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

    _excel._archive.addFile(ArchiveFile(
        'xl/worksheets/sheet$sheetNumber.xml', content.length, content));
    var _newSheet =
        _excel._archive.findFile('xl/worksheets/sheet$sheetNumber.xml');

    _newSheet!.decompress();
    var document = XmlDocument.parse(utf8.decode(_newSheet.content));
    _excel._xmlFiles['xl/worksheets/sheet$sheetNumber.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet$sheetNumber.xml';

    _excel._xmlFiles['[Content_Types].xml']
        ?.findAllElements('Types')
        .first
        .children
        .add(XmlElement(
          XmlName('Override'),
          <XmlAttribute>[
            XmlAttribute(XmlName('ContentType'),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            XmlAttribute(
                XmlName('PartName'), '/xl/worksheets/sheet$sheetNumber.xml'),
          ],
        ));
    if (_excel._xmlFiles['xl/workbook.xml'] != null) {
      _parseTable(
          _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').last);
    }
  }

  void _parseHeaderFooter(XmlElement worksheet, Sheet sheetObject) {
    final results = worksheet.findAllElements("headerFooter");
    if (results.isEmpty) return;

    final headerFooterElement = results.first;

    sheetObject.headerFooter = HeaderFooter.fromXmlElement(headerFooterElement);
  }

  void _parseColWidthsRowHeights(XmlElement worksheet, Sheet sheetObject) {
    /* parse default column width and default row height
      example XML content
      <sheetFormatPr baseColWidth="10" defaultColWidth="26.33203125" defaultRowHeight="13" x14ac:dyDescent="0.15" />
    */
    Iterable<XmlElement> results;
    results = worksheet.findAllElements("sheetFormatPr");
    if (results.isNotEmpty) {
      results.forEach((element) {
        double? defaultColWidth;
        double? defaultRowHeight;
        // default column width
        String? widthAttribute = element.getAttribute("defaultColWidth");
        if (widthAttribute != null) {
          defaultColWidth = double.tryParse(widthAttribute);
        }
        // default row height
        String? rowHeightAttribute = element.getAttribute("defaultRowHeight");
        if (rowHeightAttribute != null) {
          defaultRowHeight = double.tryParse(rowHeightAttribute);
        }

        // both values valid ?
        if (defaultColWidth != null && defaultRowHeight != null) {
          sheetObject._defaultColumnWidth = defaultColWidth;
          sheetObject._defaultRowHeight = defaultRowHeight;
        }
      });
    }

    /* parse custom column height
      example XML content
      <col min="2" max="2" width="71.83203125" customWidth="1"/>, 
      <col min="4" max="4" width="26.5" customWidth="1"/>, 
      <col min="6" max="6" width="31.33203125" customWidth="1"/>
    */
    results = worksheet.findAllElements("col");
    if (results.isNotEmpty) {
      results.forEach((element) {
        String? colAttribute =
            element.getAttribute("min"); // i think min refers to the column
        String? widthAttribute = element.getAttribute("width");
        if (colAttribute != null && widthAttribute != null) {
          int? col = int.tryParse(colAttribute);
          double? width = double.tryParse(widthAttribute);
          if (col != null && width != null) {
            col -= 1; // first col in _columnWidths is index 0
            if (col >= 0) {
              sheetObject._columnWidths[col] = width;
            }
          }
        }
      });
    }

    /* parse custom row height
      example XML content
      <row r="1" spans="1:2" ht="44" customHeight="1" x14ac:dyDescent="0.15">
    */
    results = worksheet.findAllElements("row");
    if (results.isNotEmpty) {
      results.forEach((element) {
        String? rowAttribute =
            element.getAttribute("r"); // i think min refers to the column
        String? heightAttribute = element.getAttribute("ht");
        if (rowAttribute != null && heightAttribute != null) {
          int? row = int.tryParse(rowAttribute);
          double? height = double.tryParse(heightAttribute);
          if (row != null && height != null) {
            row -= 1; // first col in _rowHeights is index 0
            if (row >= 0) {
              sheetObject._rowHeights[row] = height;
            }
          }
        }
      });
    }
  }
}
