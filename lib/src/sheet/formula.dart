part of excel;

class Formula {
  late String _formula;

  Formula._(String formula) {
    this._formula = formula;
  }

  /// Helps to initiate a custom formula
  ///```
  ///var my_custom_formula = Formula.custom('=SUM(1,2)');
  ///```
  static Formula custom(String formula) {
    return Formula._(formula);
  }

  /// get Formula
  get formula {
    return this._formula;
  }

  @override
  String toString() {
    return this._formula;
  }
}
