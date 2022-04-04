part of excel;

// For Spanning the columns and rows
// ignore: must_be_immutable
class _Span extends Equatable {
  late List<int> __start;
  late List<int> __end;

  _Span() {
    __start = <int>[];
    __end = <int>[];
  }

  set _start(List<int> val) {
    __start = val;
  }

  set _end(List<int> val) {
    __end = val;
  }

  int get rowSpanStart {
    return __start[0];
  }

  int get rowSpanEnd {
    return __end[0];
  }

  int get columnSpanStart {
    return __start[1];
  }

  int get columnSpanEnd {
    return __end[1];
  }

  @override
  List<Object?> get props => [
        __start[0],
        __start[1],
        __end[0],
        __end[1],
      ];
}
