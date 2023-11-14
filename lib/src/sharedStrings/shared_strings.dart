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
  late int count;
  _IndexingHolder(this.index, [int _count = 1]) {
    this.count = _count;
  }

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
