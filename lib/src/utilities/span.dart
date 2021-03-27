part of excel;

// For Spanning the columns and rows
class _Span extends Equatable {
  _Span();

  List<int> __start = <int>[];

  List<int> __end = <int>[];

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
        __start,
        __end,
      ];
}
