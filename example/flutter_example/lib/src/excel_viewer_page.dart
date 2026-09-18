import 'dart:io';
import 'dart:typed_data';

import 'package:excel_ce/excel_ce.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';

import 'models/sheet_model.dart';
import 'widgets/editable_data_table.dart';

class ExcelViewerPage extends StatefulWidget {
  const ExcelViewerPage({super.key});

  @override
  State<ExcelViewerPage> createState() => _ExcelViewerPageState();
}

class _ExcelViewerPageState extends State<ExcelViewerPage>
    with SingleTickerProviderStateMixin {
  List<SheetModel> _sheets = [];
  String? _fileName;
  bool _loading = false;
  TabController? _tabController;

  // ── helpers ────────────────────────────────────────────────────────────────

  String _cellString(Data? cell) {
    if (cell == null) return '';
    return switch (cell.value) {
      TextCellValue() => cell.value!.toString(),
      IntCellValue(value: final v) => v.toString(),
      DoubleCellValue(value: final v) => v.toString(),
      BoolCellValue(value: final v) => v.toString(),
      FormulaCellValue(formula: final f) => '=$f',
      DateCellValue(year: final y, month: final mo, day: final d) =>
        '$y-$mo-$d',
      DateTimeCellValue() =>
        (cell.value as DateTimeCellValue).asDateTimeLocal().toString(),
      TimeCellValue(hour: final h, minute: final m, second: final s) =>
        '$h:$m:$s',
      null => '',
    };
  }

  List<SheetModel> _parseExcel(Uint8List bytes) {
    final excel = Excel.decodeBytes(bytes);
    final result = <SheetModel>[];
    for (final name in excel.tables.keys) {
      final sheet = excel.tables[name]!;
      if (sheet.rows.isEmpty) {
        result.add(SheetModel(name: name, headers: [], rows: []));
        continue;
      }
      final headers = sheet.rows.first.map(_cellString).toList();
      final rows = sheet.rows
          .skip(1)
          .map((r) => r.map(_cellString).toList())
          .toList();
      result.add(SheetModel(name: name, headers: headers, rows: rows));
    }
    return result;
  }

  void _rebuildTabs(List<SheetModel> sheets) {
    _tabController?.dispose();
    _tabController = TabController(length: sheets.length, vsync: this);
  }

  // ── actions ────────────────────────────────────────────────────────────────

  Future<void> _pickFile() async {
    final files = await FilePicker.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx', 'xls'],
    );
    if (files.isEmpty) return;
    final picked = files.single;

    setState(() => _loading = true);
    try {
      final bytes = await picked.readAsBytes();
      final sheets = _parseExcel(bytes);
      setState(() {
        _fileName = picked.name;
        _sheets = sheets;
        _rebuildTabs(sheets);
      });
    } catch (e) {
      _showError('Failed to parse file: $e');
    } finally {
      setState(() => _loading = false);
    }
  }

  Future<void> _saveFile() async {
    if (_sheets.isEmpty) return;

    setState(() => _loading = true);
    try {
      // Rebuild Excel workbook from current sheet data
      final excel = Excel.createExcel();
      // Remove default sheet
      excel.delete('Sheet1');

      for (final model in _sheets) {
        final sheet = excel[model.name];
        // Write headers
        for (var col = 0; col < model.headers.length; col++) {
          sheet
              .cell(CellIndex.indexByColumnRow(columnIndex: col, rowIndex: 0))
              .value = TextCellValue(
            model.headers[col],
          );
        }
        // Write data rows
        for (var row = 0; row < model.rows.length; row++) {
          for (var col = 0; col < model.rows[row].length; col++) {
            sheet
                .cell(
                  CellIndex.indexByColumnRow(
                    columnIndex: col,
                    rowIndex: row + 1,
                  ),
                )
                .value = TextCellValue(
              model.rows[row][col],
            );
          }
        }
      }

      final encoded = excel.encode();
      if (encoded == null) throw Exception('Encoding failed');
      final bytes = Uint8List.fromList(encoded);

      final suggestedName = _fileName ?? 'output.xlsx';

      // Try FilePicker save dialog (desktop / Android SAF). The plugin
      // writes `bytes` to the chosen location itself and returns its Uri.
      final savedUri = await FilePicker.saveFile(
        dialogTitle: 'Save Excel file',
        fileName: suggestedName,
        allowedExtensions: ['xlsx'],
        type: FileType.custom,
        bytes: bytes,
      );

      if (savedUri != null) {
        final display = savedUri.scheme == 'file'
            ? savedUri.toFilePath()
            : savedUri.toString();
        _showSnack('Saved to $display');
      } else {
        // Fallback: save to documents directory (mobile when dialog is cancelled
        // or not supported)
        final dir = await getApplicationDocumentsDirectory();
        final file = File('${dir.path}/$suggestedName');
        await file.writeAsBytes(bytes, flush: true);
        _showSnack('Saved to ${file.path}');
      }
    } catch (e) {
      _showError('Failed to save: $e');
    } finally {
      setState(() => _loading = false);
    }
  }

  void _onCellChanged(int sheetIdx, int rowIdx, int colIdx, String value) {
    setState(() {
      final row = _sheets[sheetIdx].rows[rowIdx];
      // Extend row if needed
      while (row.length <= colIdx) {
        row.add('');
      }
      row[colIdx] = value;
    });
  }

  void _showSnack(String msg) {
    if (!mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(msg)));
  }

  void _showError(String msg) {
    if (!mounted) return;
    ScaffoldMessenger.of(
      context,
    ).showSnackBar(SnackBar(content: Text(msg), backgroundColor: Colors.red));
  }

  // ── build ──────────────────────────────────────────────────────────────────

  @override
  void dispose() {
    _tabController?.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final hasFile = _sheets.isNotEmpty;

    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(_fileName ?? 'Excel Viewer'),
        actions: [
          IconButton(
            icon: const Icon(Icons.folder_open),
            tooltip: 'Open Excel file',
            onPressed: _loading ? null : _pickFile,
          ),
          if (hasFile)
            IconButton(
              icon: const Icon(Icons.save_alt),
              tooltip: 'Save Excel file',
              onPressed: _loading ? null : _saveFile,
            ),
        ],
        bottom: hasFile && _tabController != null
            ? TabBar(
                controller: _tabController,
                isScrollable: true,
                tabs: _sheets.map((s) => Tab(text: s.name)).toList(),
              )
            : null,
      ),
      body: _loading
          ? const Center(child: CircularProgressIndicator())
          : !hasFile
          ? Center(
              child: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  Icon(
                    Icons.table_chart_outlined,
                    size: 80,
                    color: Theme.of(context).colorScheme.outline,
                  ),
                  const SizedBox(height: 16),
                  Text(
                    'Open an Excel file to get started',
                    style: Theme.of(context).textTheme.titleMedium,
                  ),
                  const SizedBox(height: 12),
                  FilledButton.icon(
                    icon: const Icon(Icons.folder_open),
                    label: const Text('Pick File'),
                    onPressed: _pickFile,
                  ),
                ],
              ),
            )
          : TabBarView(
              controller: _tabController,
              children: List.generate(
                _sheets.length,
                (i) => EditableDataTable(
                  headers: _sheets[i].headers,
                  rows: _sheets[i].rows,
                  onCellChanged: (row, col, val) =>
                      _onCellChanged(i, row, col, val),
                  frozenColumnCount: 3,
                  frozenRowCount: 3,
                  horizontalScrollPhysics: const ClampingScrollPhysics(),
                  verticalScrollPhysics: const ClampingScrollPhysics(),
                ),
              ),
            ),
    );
  }
}
