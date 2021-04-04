part of excel;

class _SharedStringsMaintainer {
  static final instance = _SharedStringsMaintainer._();

  var _map;
  var _list;
  var _index;

  factory _SharedStringsMaintainer._() {
    return _SharedStringsMaintainer();
  }

  _SharedStringsMaintainer() {
    _map = <String, _IndexingHolder>{};
    _list = <String>[];
    _index = 0;
  }

  void add(String val) {
    if (_map[val] == null) {
      _map[val] = _IndexingHolder(_index);
      _list.add(val);
      _index += 1;
    } else {
      _map[val].increaseCount();
    }
  }

  int indexOf(String val) {
    return _map[val].index;
  }

  String value(int i) {
    return i < _list.length ? _list[i] : null;
  }

  void clear() {
    _index = 0;
    _list = <String>[];
    _map = <String, _IndexingHolder>{};
  }
}

class _IndexingHolder {
  final int index;
  int count;
  _IndexingHolder(this.index, [int _count = 1]) {
    this.count = _count;
  }

  void increaseCount() {
    this.count += 1;
  }
}
