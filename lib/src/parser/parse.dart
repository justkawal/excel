part of excel;

class Parser {
  late Excel _excel;
  late List<String> _rId;
  late Map<String, String> _worksheetTargets;
  Parser._(Excel excel) {
    this._excel = excel;
    this._rId = <String>[];
    this._worksheetTargets = <String, String>{};
  }

  _startParsing() {
    _putContentXml();
    _parseRelations();
    _parseStyles(_excel._stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _parseMergedCells();
  }

  _normalizeTable(Sheet sheet) {
    if (sheet._maxRows == 0 || sheet._maxCols == 0) {
      sheet._sheetData.clear();
    }
    sheet._countRowAndCol();
  }

  _putContentXml() {
    var file = _excel._archive.findFile("[Content_Types].xml");

    if (file == null) {
      _damagedExcel();
    }
    file!.decompress();
    _excel._xmlFiles["[Content_Types].xml"] =
        XmlDocument.parse(utf8.decode(file.content));
  }

  _parseRelations() {
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

  _parseSharedStrings() {
    var sharedStrings =
        _excel._archive.findFile('xl/${_excel._sharedStringsTarget}');
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
      _excel._archive.addFile(ArchiveFile(
          'xl/${_excel._sharedStringsTarget}', content.length, content));
      sharedStrings =
          _excel._archive.findFile('xl/${_excel._sharedStringsTarget}');
    }
    sharedStrings!.decompress();
    var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
    _excel._xmlFiles["xl/${_excel._sharedStringsTarget}"] = document;

    document.findAllElements('si').forEach((node) {
      _parseSharedString(node);
    });
  }

  _parseSharedString(XmlElement node) {
    final sharedString = SharedString(node: node);
    _excel._sharedStrings.add(sharedString);
  }

  _parseContent({bool run = true}) {
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

  _parseMergedCells() {
    Map spannedCells = <String, List<String>>{};
    _excel._sheets.forEach((sheetName, node) {
      _excel._availSheet(sheetName);
      XmlElement elementNode = node as XmlElement;
      List spanList = <String>[];

      elementNode.findAllElements('mergeCell').forEach((elemen) {
        String? ref = elemen.getAttribute('ref');
        if (ref != null && ref.contains(':') && ref.split(':').length == 2) {
          if (!_excel._sheetMap[sheetName]!._spannedItems.contains(ref)) {
            _excel._sheetMap[sheetName]!._spannedItems.add(ref);
          }

          String startCell = ref.split(':')[0], endCell = ref.split(':')[1];

          if (!spanList.contains(startCell)) {
            spanList.add(startCell);
          }
          spannedCells[sheetName] = spanList;

          List<int> startIndex = _cellCoordsFromCellId(startCell),
              endIndex = _cellCoordsFromCellId(endCell);
          _Span spanObj = _Span();
          spanObj._start = [startIndex[0], startIndex[1]];
          spanObj._end = [endIndex[0], endIndex[1]];
          if (!_excel._sheetMap[sheetName]!._spanList.contains(spanObj)) {
            _excel._sheetMap[sheetName]!._spanList.add(spanObj);
          }
          _excel._mergeChangeLookup = sheetName;
        }
      });
    });

    // Remove those cells which are present inside the
    _excel._sheetMap.forEach((sheetName, sheetObject) {
      if (spannedCells.containsKey(sheetName)) {
        sheetObject._sheetData.forEach((row, colMap) {
          colMap.forEach((col, dataObject) {
            if (!(spannedCells[sheetName].contains(getCellId(col, row)))) {
              _excel[sheetName]._sheetData[row]?.remove(col);
            }
          });
        });
      }
    });
  }

  // Reading the styles from the excel file.
  _parseStyles(String _stylesTarget) {
    var styles = _excel._archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));
      _excel._xmlFiles['xl/$_stylesTarget'] = document;

      _excel._fontStyleList = <_FontStyle>[];
      _excel._patternFill = <String>[];
      _excel._cellStyleList = <CellStyle>[];

      Iterable<XmlElement> fontList = document.findAllElements('font');

      document.findAllElements('patternFill').forEach((node) {
        String patternType = node.getAttribute('patternType').toString(), rgb;
        if (node.children.isNotEmpty) {
          node.findElements('fgColor').forEach((child) {
            rgb = node.getAttribute('rgb').toString();
            _excel._patternFill.add(rgb);
          });
        } else {
          _excel._patternFill.add(patternType);
        }
      });

      document.findAllElements('cellXfs').forEach((node1) {
        node1.findAllElements('xf').forEach((node) {
          _excel._numFormats.add(_getFontIndex(node, 'numFmtId'));

          String fontColor = "FF000000", backgroundColor = "none";
          String? fontFamily;

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
            var _single_underline = _nodeChildren(font, 'u');
            if (_single_underline != null) {
              underline = Underline.Single;
            }

            /// Checking for font Family
            var _family = _nodeChildren(font, 'name', attribute: 'val');
            if (_family != null && _family != true) {
              fontFamily = _family;
            }

            _fontStyle.isBold = isBold;
            _fontStyle.isItalic = isItalic;
            _fontStyle.fontSize = fontSize;
            _fontStyle.fontFamily = fontFamily;
            _fontStyle._fontColorHex = fontColor;
          }

          /// If `-1` is returned then it indicates that `_fontStyle` is not present in the `_fontStyleList`
          if (_fontStyleIndex(_excel._fontStyleList, _fontStyle) == -1) {
            _excel._fontStyleList.add(_fontStyle);
          }

          int fillId = _getFontIndex(node, 'fillId');
          if (fillId < _excel._patternFill.length) {
            backgroundColor = _excel._patternFill[fillId];
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

          CellStyle cellStyle = CellStyle(
            fontColorHex: fontColor,
            fontFamily: fontFamily,
            fontSize: fontSize,
            bold: isBold,
            italic: isItalic,
            underline: underline,
            backgroundColorHex: backgroundColor,
            horizontalAlign: horizontalAlign,
            verticalAlign: verticalAlign,
            textWrapping: textWrapping,
            rotation: rotation,
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

  _parseTable(XmlElement node) {
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

  _parseCell(XmlElement node, Sheet sheetObject, int rowIndex, String name) {
    int? colIndex = _getCellNumber(node);
    if (colIndex == null) {
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

    if (node.children.isEmpty) {
      return;
    }

    var value, type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        value = _excel._sharedStrings
            .value(int.parse(_parseValue(node.findElements('v').first)));
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
        var formulaNode = node.findElements('f');
        var content = valueNode.first;
        if (formulaNode.isNotEmpty) {
          value = Formula.custom(_parseValue(formulaNode.first).toString());
        } else {
          if (s1 != null) {
            var fmtId = _excel._numFormats[s];
            // date
            if (((fmtId >= 14) && (fmtId <= 17)) ||
                (fmtId == 22) ||
                (fmtId == 164)) {
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
    }
    sheetObject.updateCell(
        CellIndex.indexByColumnRow(columnIndex: colIndex, rowIndex: rowIndex),
        value,
        cellStyle: _excel._cellStyleList[s]);
  }

  static _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
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
  _createSheet(String newSheet) {
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
          XmlAttribute(XmlName('Target'), 'worksheets/sheet${sheetNumber}.xml'),
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
            XmlAttribute(XmlName('sheetId'), '${sheetNumber}'),
            XmlAttribute(XmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] = 'worksheets/sheet${sheetNumber}.xml';

    var content = utf8.encode(
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

    _excel._archive.addFile(ArchiveFile(
        'xl/worksheets/sheet${sheetNumber}.xml', content.length, content));
    var _newSheet =
        _excel._archive.findFile('xl/worksheets/sheet${sheetNumber}.xml');

    _newSheet!.decompress();
    var document = XmlDocument.parse(utf8.decode(_newSheet.content));
    _excel._xmlFiles['xl/worksheets/sheet${sheetNumber}.xml'] = document;
    _excel._xmlSheetId[newSheet] = 'xl/worksheets/sheet${sheetNumber}.xml';

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
                XmlName('PartName'), '/xl/worksheets/sheet${sheetNumber}.xml'),
          ],
        ));
    if (_excel._xmlFiles['xl/workbook.xml'] != null)
      _parseTable(
          _excel._xmlFiles['xl/workbook.xml']!.findAllElements('sheet').last);
  }

  void _parseHeaderFooter(XmlElement worksheet, Sheet sheetObject) {
    final results = worksheet.findAllElements("headerFooter");
    if (results.isEmpty) return;

    final headerFooterElement = results.first;

    sheetObject.headerFooter = HeaderFooter.fromXmlElement(headerFooterElement);
  }
}
