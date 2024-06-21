part of excel;

class _SharedStringsMaintainer {
  final Map<SharedString, _IndexingHolder> _map =
      <SharedString, _IndexingHolder>{};
  final Map<String, SharedString> _mapString = <String, SharedString>{};
  final List<SharedString> _list = <SharedString>[];
  int _index = 0;

  _SharedStringsMaintainer._();

  SharedString? tryFind(String val) {
    return _mapString[val];
  }

  SharedString addFromString(String val) {
    final newSharedString = SharedString(
        node: XmlElement(XmlName('si'), [], [
      XmlElement(XmlName('t'),
          [XmlAttribute(XmlName("space", "xml"), "preserve")], [XmlText(val)]),
    ]));

    add(newSharedString, val);
    return newSharedString;
  }

  void add(SharedString val, String key) {
    _map[val]?.increaseCount();
    _map.putIfAbsent(val, () {
      _mapString[key] = val;
      _list.add(val);
      return _IndexingHolder(_index++);
    });
  }

  int indexOf(SharedString val) {
    return _map[val] != null ? _map[val]!.index : -1;
  }

  SharedString? value(int i) {
    if (i < _list.length) {
      return _list[i];
    } else {
      return null;
    }
  }

  void clear() {
    _index = 0;
    _list.clear();
    _map.clear();
    _mapString.clear();
  }
}

class _IndexingHolder {
  final int index;
  int count;

  _IndexingHolder(this.index, [int _count = 1]) : count = _count;

  void increaseCount() {
    this.count += 1;
  }
}

class SharedString {
  final XmlElement node;
  final int _hashCode;

  SharedString({required this.node}) : _hashCode = node.toString().hashCode;

  @override
  String toString() {
    assert(false,
        'prefer stringValue over SharedString.toString() in development');
    return stringValue;
  }

  TextSpan get textSpan {
    bool getBool(XmlElement element) {
      return bool.tryParse(element.getAttribute('val') ?? '') ?? true;
    }

    int getDouble(XmlElement element) {
      // Should be double
      return double.parse(element.getAttribute('val')!).toInt();
    }

    String? text;
    List<TextSpan>? children;

    /// SharedStringItem
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem?view=openxml-3.0.1
    assert(node.localName == 'si'); //18.4.8 si (String Item)

    for (final child in node.childElements) {
      switch (child.localName) {
        /// Text
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1
        case 't': //18.4.12 t (Text)
          text = (text ?? '') + child.innerText;
          break;

        /// Rich Text Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.run?view=openxml-3.0.1
        case 'r': //18.4.4 r (Rich Text Run)
          var style = CellStyle();
          for (final runChild in child.childElements) {
            switch (runChild.localName) {
              /// RunProperties
              /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.runproperties?view=openxml-3.0.1
              case 'rPr':
                for (final runProperty in runChild.childElements) {
                  switch (runProperty.localName) {
                    case 'b': //18.8.2 b (Bold)
                      style = style.copyWith(boldVal: getBool(runProperty));
                      break;
                    case 'i': //18.8.26 i (Italic)
                      style = style.copyWith(italicVal: getBool(runProperty));
                      break;
                    case 'u': //18.4.13 u (Underline)
                      style = style.copyWith(
                          underlineVal:
                              runProperty.getAttribute('val') == 'double'
                                  ? Underline.Double
                                  : Underline.Single);
                      break;
                    case 'sz': //18.4.11 sz (Font Size)
                      style =
                          style.copyWith(fontSizeVal: getDouble(runProperty));
                      break;
                    case 'rFont': //18.4.5 rFont (Font)
                      style = style.copyWith(
                          fontFamilyVal: runProperty.getAttribute('val'));
                      break;
                    case 'color': //18.3.1.15 color (Data Bar Color)
                      style = style.copyWith(
                          fontColorHexVal:
                              runProperty.getAttribute('rgb')?.excelColor);
                      break;
                  }
                }
                break;

              /// Text
              case 't': //18.4.12 t (Text)
                if (children == null) children = [];
                children.add(TextSpan(text: runChild.innerText, style: style));
                break;
            }
          }
          break;

        /// Phonetic Run
        /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.phoneticrun?view=openxml-3.0.1
        case 'rPh': //18.4.6 rPh (Phonetic Run)
          break;
      }
    }

    return TextSpan(text: text, children: children);
  }

  String get stringValue {
    var buffer = StringBuffer();
    node.findAllElements('t').forEach((child) {
      if (child.parentElement == null ||
          child.parentElement!.name.local != 'rPh') {
        buffer.write(Parser._parseValue(child));
      }
    });
    return buffer.toString();
  }

  @override
  int get hashCode => _hashCode;

  @override
  operator ==(Object other) {
    return other is SharedString &&
        other.hashCode == _hashCode &&
        other.stringValue == stringValue;
  }

  bool matches(String value) {
    return value.isNotEmpty && value == stringValue;
  }
}

class TextSpan {
  final String? text;
  final List<TextSpan>? children;
  final CellStyle? style;

  const TextSpan({this.children, this.text, this.style});

  @override
  String toString() {
    String r = '';
    if (text != null) r += text!;
    if (children != null) r += children!.join();
    return r;
  }

  @override
  operator ==(Object other) {
    if (identical(this, other)) return true;
    if (other.runtimeType != runtimeType) return false;
    return other is TextSpan &&
        other.text == text &&
        other.style == style &&
        ListEquality().equals(other.children, children);
  }

  @override
  int get hashCode =>
      Object.hash(text, style, Object.hashAll(children ?? const []));
}
