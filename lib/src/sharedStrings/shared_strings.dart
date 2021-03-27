part of excel;

class _TempDataHolder {
  final int index;
  late int count;
  _TempDataHolder(this.index, [int count = 1]) {
    this.count = count;
  }

  void increaseCount() {
    this.count += 1;
  }
}

class _OptimizedList<T> {
  var _map = Map<T, _TempDataHolder>();
  var _list = <T>[];
  var _index = 0;

  void add(T val) {
    if (_map[val] == null) {
      _map[val] = _TempDataHolder(_index);
      _list.add(val);
      _index += 1;
    } else {
      _map[val]!.increaseCount();
    }
  }

  int indexOf(T val) {
    return _map[val] != null ? _map[val]!.index : -1;
  }

  T? value(int index) {
    return index < _list.length ? _list[index] : null;
  }

  void clear() {
    _index = 0;
    _list = <T>[];
    _map = <T, _TempDataHolder>{};
  }
}

class _SharedStringsMaintainer {
  static final instance = _SharedStringsMaintainer._();

  var list;

  _SharedStringsMaintainer() {
    list = _OptimizedList<String>();
  }

  factory _SharedStringsMaintainer._() {
    return _SharedStringsMaintainer();
  }
}
