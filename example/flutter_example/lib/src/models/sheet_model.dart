class SheetModel {
  final String name;
  final List<String> headers;
  final List<List<String>> rows;

  SheetModel({
    required this.name,
    required this.headers,
    required this.rows,
  });
}
