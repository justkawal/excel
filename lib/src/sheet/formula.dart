part of excel;

class Formula {
  final String formula;

  Formula._(this.formula);

  /// Helps to initiate a custom formula
  ///```
  ///var my_custom_formula = Formula.custom('=SUM(1,2)');
  ///```
  static Formula custom(String formula) {
    return Formula._(formula);
  }

  @override
  String toString() {
    return formula;
  }
}
