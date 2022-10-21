part of excel_facility;

class _SharedStringsMaintainer {
  static final instance = _SharedStringsMaintainer._();

  late Map<String, _IndexingHolder> _map;
  late List<String> _list;
  late int _index;

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
      _map[val]!.increaseCount();
    }
  }

  int indexOf(String val) {
    return _map[val] != null ? _map[val]!.index : -1;
  }

  String? value(int i) {
    return i < _list.length ? _list[i] : null;
  }

  void clear() {
    _index = 0;
    _list = <String>[];
    _map = <String, _IndexingHolder>{};
  }

  void ensureReinitialize() {
    _map = <String, _IndexingHolder>{};
    _list = <String>[];
    _index = 0;
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
