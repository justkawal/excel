import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';
import 'package:flutter_example/src/widgets/editable_data_table.dart';

// ── Helpers ───────────────────────────────────────────────────────────────────

List<String> _makeHeaders(int cols) =>
    List.generate(cols, (i) => 'Col_$i');

List<List<String>> _makeRows(int rows, int cols) =>
    List.generate(rows, (r) => List.generate(cols, (c) => 'R${r}C$c'));

Widget _wrap(Widget child) =>
    MaterialApp(home: Scaffold(body: child));

// ── Tests ─────────────────────────────────────────────────────────────────────

void main() {
  // ── 1. Empty state ──────────────────────────────────────────────────────────

  group('EditableDataTable – empty state', () {
    testWidgets('shows "Sheet is empty" when headers and rows are both empty',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const [],
        rows: const [],
        onCellChanged: (_, __, ___) {},
      )));
      expect(find.text('Sheet is empty'), findsOneWidget);
    });
  });

  // ── 2. Rendering ────────────────────────────────────────────────────────────

  group('EditableDataTable – rendering', () {
    testWidgets('renders all header labels', (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name', 'Age', 'Email'],
        rows: const [
          ['Alice', '30', 'alice@example.com']
        ],
        onCellChanged: (_, __, ___) {},
      )));
      expect(find.text('Name'), findsOneWidget);
      expect(find.text('Age'), findsOneWidget);
      expect(find.text('Email'), findsOneWidget);
    });

    testWidgets('renders visible cell values', (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['X', 'Y'],
        rows: const [
          ['hello', 'world'],
          ['foo', 'bar'],
        ],
        onCellChanged: (_, __, ___) {},
      )));
      expect(find.text('hello'), findsOneWidget);
      expect(find.text('world'), findsOneWidget);
    });

    testWidgets('rows shorter than headers render without throwing',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['A', 'B', 'C'],
        rows: const [
          ['only-a']
        ],
        onCellChanged: (_, __, ___) {},
      )));
      expect(find.text('only-a'), findsOneWidget);
      expect(tester.takeException(), isNull);
    });
  });

  // ── 3. Edit dialog ──────────────────────────────────────────────────────────

  group('EditableDataTable – edit dialog', () {
    testWidgets('tapping a cell opens the AlertDialog', (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name'],
        rows: const [
          ['Alice']
        ],
        onCellChanged: (_, __, ___) {},
      )));
      await tester.tap(find.text('Alice'));
      await tester.pumpAndSettle();
      expect(find.byType(AlertDialog), findsOneWidget);
    });

    testWidgets('dialog pre-fills the current cell value', (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name'],
        rows: const [
          ['Alice']
        ],
        onCellChanged: (_, __, ___) {},
      )));
      await tester.tap(find.text('Alice'));
      await tester.pumpAndSettle();
      final tf = tester.widget<TextField>(find.byType(TextField));
      expect(tf.controller?.text, 'Alice');
    });

    testWidgets('dialog title contains the column label and row number',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Score'],
        rows: const [
          ['42']
        ],
        onCellChanged: (_, __, ___) {},
      )));
      await tester.tap(find.text('42'));
      await tester.pumpAndSettle();
      // title is "Score — Row 1"
      expect(find.textContaining('Score'), findsWidgets);
      expect(find.textContaining('Row 1'), findsOneWidget);
    });

    testWidgets('confirming edit calls onCellChanged with correct args',
        (tester) async {
      int? cbRow, cbCol;
      String? cbVal;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name'],
        rows: const [
          ['Alice']
        ],
        onCellChanged: (r, c, v) {
          cbRow = r;
          cbCol = c;
          cbVal = v;
        },
      )));
      await tester.tap(find.text('Alice'));
      await tester.pumpAndSettle();
      await tester.enterText(find.byType(TextField), 'Bob');
      await tester.tap(find.text('Save'));
      await tester.pumpAndSettle();
      expect(cbRow, 0);
      expect(cbCol, 0);
      expect(cbVal, 'Bob');
    });

    testWidgets('cancelling does NOT call onCellChanged', (tester) async {
      var called = false;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name'],
        rows: const [
          ['Alice']
        ],
        onCellChanged: (_, __, ___) => called = true,
      )));
      await tester.tap(find.text('Alice'));
      await tester.pumpAndSettle();
      await tester.tap(find.text('Cancel'));
      await tester.pumpAndSettle();
      expect(called, isFalse);
    });

    testWidgets('correct row / col indices for a non-first cell',
        (tester) async {
      int? cbRow, cbCol;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['A', 'B', 'C'],
        rows: const [
          ['r0c0', 'r0c1', 'r0c2'],
          ['r1c0', 'r1c1', 'r1c2'],
        ],
        onCellChanged: (r, c, v) {
          cbRow = r;
          cbCol = c;
        },
      )));
      await tester.tap(find.text('r1c2'));
      await tester.pumpAndSettle();
      await tester.tap(find.text('Save'));
      await tester.pumpAndSettle();
      expect(cbRow, 1);
      expect(cbCol, 2);
    });
  });

  // ── 4. Large data ───────────────────────────────────────────────────────────

  group('EditableDataTable – large data performance', () {
    testWidgets('5 000 rows × 100 cols pumps without throwing', (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: _makeHeaders(100),
        rows: _makeRows(5000, 100),
        onCellChanged: (_, __, ___) {},
      )));
      await tester.pump();
      expect(find.text('Col_0'), findsOneWidget);
      expect(find.text('R0C0'), findsOneWidget);
      expect(tester.takeException(), isNull);
    });

    testWidgets('row 500 is NOT in the widget tree (virtualisation check)',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: _makeHeaders(5),
        rows: _makeRows(5000, 5),
        onCellChanged: (_, __, ___) {},
        rowHeight: 48,
      )));
      await tester.pump();
      // SliverFixedExtentList should not build rows outside the viewport
      expect(find.text('R500C0'), findsNothing);
    });

    testWidgets('editing in a large dataset fires the correct callback',
        (tester) async {
      String? editedVal;
      int? editedRow, editedCol;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: _makeHeaders(5),
        rows: _makeRows(100, 5),
        onCellChanged: (r, c, v) {
          editedRow = r;
          editedCol = c;
          editedVal = v;
        },
      )));
      await tester.pump();
      await tester.tap(find.text('R0C0'));
      await tester.pumpAndSettle();
      await tester.enterText(find.byType(TextField), 'Updated');
      await tester.tap(find.text('Save'));
      await tester.pumpAndSettle();
      expect(editedRow, 0);
      expect(editedCol, 0);
      expect(editedVal, 'Updated');
    });
  });

  // ── 5. Styling ──────────────────────────────────────────────────────────────

  group('EditableDataTable – styling', () {
    testWidgets('accepts a custom EditableDataTableStyle without throwing',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['A'],
        rows: const [
          ['1']
        ],
        onCellChanged: (_, __, ___) {},
        style: const EditableDataTableStyle(
          headerBackgroundColor: Colors.blue,
          rowEvenColor: Colors.white,
          rowOddColor: Color(0xFFF5F5F5),
          borderColor: Colors.grey,
          headerTextStyle:
              TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
          cellTextStyle: TextStyle(fontSize: 13),
        ),
      )));
      expect(tester.takeException(), isNull);
      expect(find.text('A'), findsOneWidget);
    });

    testWidgets('custom columnWidth and rowHeight are used in header SizedBox',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name'],
        rows: const [
          ['Alice']
        ],
        onCellChanged: (_, __, ___) {},
        columnWidth: 200,
        rowHeight: 60,
      )));
      // Find the SizedBox wrapping the header cell text
      final sizedBoxes = tester
          .widgetList<SizedBox>(
            find.ancestor(
              of: find.text('Name'),
              matching: find.byType(SizedBox),
            ),
          )
          .toList();
      expect(
        sizedBoxes.any((b) => b.width == 200 && b.height == 60),
        isTrue,
      );
    });
  });

  // ── 6. Sorting ──────────────────────────────────────────────────────────────

  group('EditableDataTable – sorting', () {
    // Shared dataset: original order Charlie(0), Alice(1), Bob(2)
    const sortHeaders = ['Name', 'Score'];
    const sortRows = [
      ['Charlie', '90'],
      ['Alice', '60'],
      ['Bob', '80'],
    ];

    testWidgets('tapping a header sorts column ascending', (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: sortHeaders,
        rows: sortRows,
        onCellChanged: (_, __, ___) {},
      )));
      // All rows present before sort
      expect(find.text('Charlie'), findsOneWidget);
      expect(find.text('Alice'), findsOneWidget);
      expect(find.text('Bob'), findsOneWidget);

      // Tap 'Name' header → ascending sort (Alice < Bob < Charlie)
      await tester.tap(find.text('Name'));
      await tester.pumpAndSettle();

      // All rows still rendered
      expect(find.text('Alice'), findsOneWidget);
      expect(find.text('Bob'), findsOneWidget);
      expect(find.text('Charlie'), findsOneWidget);

      // Alice must appear above Charlie in the viewport
      final aliceY = tester.getTopLeft(find.text('Alice')).dy;
      final charlieY = tester.getTopLeft(find.text('Charlie')).dy;
      expect(aliceY, lessThan(charlieY));
    });

    testWidgets('tapping a header twice sorts column descending',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: sortHeaders,
        rows: sortRows,
        onCellChanged: (_, __, ___) {},
      )));
      await tester.tap(find.text('Name')); // ascending
      await tester.pumpAndSettle();
      await tester.tap(find.text('Name')); // descending
      await tester.pumpAndSettle();

      // Descending: Charlie > Bob > Alice → Charlie should appear above Alice
      final charlieY = tester.getTopLeft(find.text('Charlie')).dy;
      final aliceY = tester.getTopLeft(find.text('Alice')).dy;
      expect(charlieY, lessThan(aliceY));
    });

    testWidgets('tapping a header three times restores original order',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: sortHeaders,
        rows: sortRows,
        onCellChanged: (_, __, ___) {},
      )));
      for (var i = 0; i < 3; i++) {
        await tester.tap(find.text('Name'));
        await tester.pumpAndSettle();
      }
      // Original order: Charlie(0), Alice(1), Bob(2)
      final charlieY = tester.getTopLeft(find.text('Charlie')).dy;
      final aliceY = tester.getTopLeft(find.text('Alice')).dy;
      final bobY = tester.getTopLeft(find.text('Bob')).dy;
      expect(charlieY, lessThan(aliceY));
      expect(aliceY, lessThan(bobY));
    });

    testWidgets('onCellChanged reports original row index when sorted',
        (tester) async {
      int? cbRow, cbCol;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: sortHeaders,
        rows: sortRows,
        onCellChanged: (r, c, v) {
          cbRow = r;
          cbCol = c;
        },
      )));
      // Sort ascending by Name → visual[0] = 'Alice' (original index 1)
      await tester.tap(find.text('Name'));
      await tester.pumpAndSettle();

      // Tap 'Alice' — visually row 0 but originally row 1
      await tester.tap(find.text('Alice'));
      await tester.pumpAndSettle();
      await tester.tap(find.text('Save'));
      await tester.pumpAndSettle();

      expect(cbRow, 1); // original index, NOT visual index
      expect(cbCol, 0);
    });

    testWidgets('sort indicator icon appears on the active sort column',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: sortHeaders,
        rows: sortRows,
        onCellChanged: (_, __, ___) {},
      )));
      // Unsorted → no arrow
      expect(find.byIcon(Icons.arrow_upward), findsNothing);

      // 1st tap → ascending → arrow visible
      await tester.tap(find.text('Name'));
      await tester.pumpAndSettle();
      expect(find.byIcon(Icons.arrow_upward), findsOneWidget);

      // 3rd tap → clear sort → arrow gone
      await tester.tap(find.text('Name')); // descending
      await tester.pumpAndSettle();
      await tester.tap(find.text('Name')); // clear
      await tester.pumpAndSettle();
      expect(find.byIcon(Icons.arrow_upward), findsNothing);
    });

    testWidgets('switching sort column clears previous indicator',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: sortHeaders,
        rows: sortRows,
        onCellChanged: (_, __, ___) {},
      )));
      await tester.tap(find.text('Name')); // sort by Name
      await tester.pumpAndSettle();
      expect(find.byIcon(Icons.arrow_upward), findsOneWidget);

      await tester.tap(find.text('Score')); // switch sort to Score
      await tester.pumpAndSettle();
      // Still one indicator — now on Score column, not two
      expect(find.byIcon(Icons.arrow_upward), findsOneWidget);
    });
  });

  // ── 7. Frozen rows / columns ─────────────────────────────────────────────────

  group('EditableDataTable – frozen rows/columns', () {
    const frozenHeaders = ['ID', 'Name', 'Score', 'Grade'];
    const frozenRows = [
      ['1', 'Alice', '95', 'A'],
      ['2', 'Bob', '82', 'B'],
      ['3', 'Charlie', '74', 'C'],
      ['4', 'Diana', '88', 'B+'],
    ];

    testWidgets('frozenRowCount clamps gracefully when larger than rows.length',
        (tester) async {
      // 2 data rows, frozenRowCount=10 → should not throw
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: frozenHeaders,
        rows: const [
          ['1', 'Alice', '95', 'A'],
          ['2', 'Bob', '82', 'B'],
        ],
        onCellChanged: (_, __, ___) {},
        frozenRowCount: 10,
      )));
      await tester.pump();
      expect(tester.takeException(), isNull);
      expect(find.text('Alice'), findsOneWidget);
    });

    testWidgets('frozenColumnCount clamps gracefully when larger than headers',
        (tester) async {
      // 4 headers, frozenColumnCount=100 → should not throw
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: frozenHeaders,
        rows: frozenRows,
        onCellChanged: (_, __, ___) {},
        frozenColumnCount: 100,
      )));
      await tester.pump();
      expect(tester.takeException(), isNull);
    });

    testWidgets('frozenColumnCount: 0 renders single-pane normally',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: frozenHeaders,
        rows: frozenRows,
        onCellChanged: (_, __, ___) {},
        frozenColumnCount: 0,
      )));
      await tester.pump();
      expect(find.text('ID'), findsOneWidget);
      expect(find.text('Name'), findsOneWidget);
      expect(find.text('Alice'), findsOneWidget);
      expect(tester.takeException(), isNull);
    });

    testWidgets('frozenColumnCount: 1 renders dual-pane without error',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: frozenHeaders,
        rows: frozenRows,
        onCellChanged: (_, __, ___) {},
        frozenColumnCount: 1,
      )));
      await tester.pump();
      // Header labels in the left (frozen) pane — 'ID' appears in both panes
      expect(find.text('ID'), findsWidgets);
      expect(tester.takeException(), isNull);
    });

    testWidgets('frozenColumnCount: 1 — frozen cell is tappable and fires callback',
        (tester) async {
      int? cbRow, cbCol;
      String? cbVal;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: frozenHeaders,
        rows: frozenRows,
        onCellChanged: (r, c, v) {
          cbRow = r;
          cbCol = c;
          cbVal = v;
        },
        frozenColumnCount: 1,
      )));
      await tester.pump();

      // Tap the first 'Alice' found — it can be in the right pane; 'ID' frozen col
      // Tap cell '1' (ID column, first row) which is in the frozen left pane
      await tester.tap(find.text('1').first);
      await tester.pumpAndSettle();
      expect(find.byType(AlertDialog), findsOneWidget);

      await tester.enterText(find.byType(TextField), '99');
      await tester.tap(find.text('Save'));
      await tester.pumpAndSettle();

      expect(cbRow, 0);
      expect(cbCol, 0);
      expect(cbVal, '99');
    });

    testWidgets(
        'frozenRowCount: 1 combined with frozenColumnCount: 1 pumps without error',
        (tester) async {
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: frozenHeaders,
        rows: frozenRows,
        onCellChanged: (_, __, ___) {},
        frozenRowCount: 1,
        frozenColumnCount: 1,
      )));
      await tester.pump();
      expect(tester.takeException(), isNull);
      // Header row visible
      expect(find.text('ID'), findsWidgets);
    });

    testWidgets('onCellChanged reports original row index when frozen + sorted',
        (tester) async {
      int? cbRow, cbCol;
      await tester.pumpWidget(_wrap(EditableDataTable(
        headers: const ['Name', 'Score'],
        rows: const [
          ['Charlie', '90'],
          ['Alice', '60'],
          ['Bob', '80'],
        ],
        onCellChanged: (r, c, v) {
          cbRow = r;
          cbCol = c;
        },
        frozenColumnCount: 1,
      )));

      // Sort ascending by Name → visual[0] = 'Alice' (original index 1)
      await tester.tap(find.text('Name').first);
      await tester.pumpAndSettle();

      // 'Alice' appears in the frozen left pane; tap it
      await tester.tap(find.text('Alice').first);
      await tester.pumpAndSettle();
      await tester.tap(find.text('Save'));
      await tester.pumpAndSettle();

      // Original index is 1 regardless of frozen pane or sort order
      expect(cbRow, 1);
      expect(cbCol, 0);
    });
  });
}
