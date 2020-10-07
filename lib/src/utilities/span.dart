part of excel;

// For Spanning the columns and rows
class _Span {
  _Span();

  List<int> __start = List<int>();

  List<int> __end = List<int>();

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

  ///
  ///returns true if the two objects are same
  ///
  @override
  bool operator ==(o) {
    return this.rowSpanStart == o.rowSpanStart &&
        this.rowSpanEnd == o.rowSpanEnd &&
        this.columnSpanStart == o.columnSpanStart &&
        this.columnSpanEnd == o.columnSpanEnd;
  }
}
