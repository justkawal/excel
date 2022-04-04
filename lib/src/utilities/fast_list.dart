part of excel;

// A helper class to optimized the usage of Maps
class FastList<K> {
  late Map<K, int> _map;
  late int _index;

  FastList() {
    _map = <K, int>{};
    _index = 0;
  }

  FastList.from(FastList<K> other) {
    _map = Map<K, int>.from(other._map);
    _index = other._index;
  }

  void add(K key) {
    if (_map[key] == null) {
      _map[key] = _index;
      _index += 1;
    }
  }

  bool contains(K key) {
    return _map[key] != null;
  }

  void remove(K key) {
    _map.remove(key);
  }

  void clear() {
    _index = 0;
    _map = <K, int>{};
  }

  List<K> get keys => _map.keys.toList();

  bool get isNotEmpty => _map.isNotEmpty;
}
