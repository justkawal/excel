class Relations {
  final Map<String, String> _targets = {};

  String? targetById(String id) => _targets[id];

  @override
  String toString() {
    return 'Relations{_targets: $_targets}';
  }
}

class RelationsByFile {
  final Map<String, Relations> _relationsByFiles = {};

  void addTarget(String fileName, String id, String target) {
    final relationFile = _relationsByFiles[fileName] ??= Relations();
    relationFile._targets[id] = target;
  }

  Relations? relations(String fileName) {
    return _relationsByFiles[fileName];
  }

  String? target(String fileName, String id) {
    return _relationsByFiles[fileName]?.targetById(id);
  }

  @override
  String toString() {
    return 'RelationsByFile{_relations: $_relationsByFiles}';
  }
}
