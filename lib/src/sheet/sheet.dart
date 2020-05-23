part of excel;

class Sheet {
  Excel _excel;
  Row _row;
  Sheet._(Excel excel) {
    this._excel = excel;
  }

  /// Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  /// Add the sheet details in the workbook.xml. as well as in the workbook.xml.rels
  /// Then add the sheet physically into the [_xmlFiles] so as to get it into the archieve.
  /// Also add it into the [_sheets] and [_tables] map so as to allow the editing.
  Sheet get _newSheet {
      List<XmlNode> list =
          _excel._xmlFiles['xl/workbook.xml'].findAllElements('sheets').first.children;
      if (list.isEmpty) {
        throw ArgumentError('');
      }
      int _sheetId = -1;
      List<int> sheetIdList = List<int>();

      _excel._xmlFiles['xl/workbook.xml']
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

      _excel._xmlFiles['xl/_rels/workbook.xml.rels']
          .findAllElements('Relationships')
          .first
          .children
          .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
            XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
            XmlAttribute(XmlName('Type'), '$_relationships/worksheet'),
            XmlAttribute(
                XmlName('Target'), 'worksheets/sheet${sheetNumber + 1}.xml'),
          ]));

      _excel._xmlFiles['xl/workbook.xml']
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

      _excel._worksheetTargets['rId$ridNumber'] =
          'worksheets/sheet${sheetNumber + 1}.xml';

      var content = utf8.encode(
          "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac xr xr2 xr3\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\"> <dimension ref=\"A1\"/> <sheetViews> <sheetView tabSelected=\"1\" workbookViewId=\"0\"/> </sheetViews> <sheetData/> <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/> </worksheet>");

      _excel._archive.addFile(ArchiveFile('xl/worksheets/sheet${sheetNumber + 1}.xml',
          content.length, content));
      var _newSheet = _excel._archive.findFile('xl/${_excel._sharedStringsTarget}');

      _newSheet.decompress();
      var document = parse(utf8.decode(_newSheet.content));
      if (_excel._xmlFiles != null) {
        _excel._xmlFiles['xl/worksheets/sheet${sheetNumber + 1}.xml'] = document;
      }

      _excel._xmlFiles['[Content_Types].xml']
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
      _parseTable(_excel._xmlFiles['xl/workbook.xml'].findAllElements('sheet').last);
    }
  }
}
