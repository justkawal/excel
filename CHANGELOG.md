# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [4.0.4]
- Fixed Bug that Merge Cells were not parsed

## [4.0.3]
- Fix Archive version

## [4.0.2] - 2023-12-23

### Modifications

- Modified Readme.md

## [4.0.1] - 2023-12-23

### Modifications

- Modified Readme.md

### Breaking Changes

- `CellIndex.indexByColumnRow()` now requires non-null integers of row index and column index

## [4.0.0] - 2023-11-25

### Breaking Changes

- Renamed `Formula` to `FormulaCellValue`
- Cells value now represented by the sealed class `CellValue` instead of `dynamic`. Subtypes are `TextCellValue` `FormulaCellValue`, `IntCellValue`, `DoubleCellValue`, `DateCellValue`, `TextCellValue`, `BoolCellValue`, `TimeCellValue`, `DateTimeCellValue` and they allow for exhaustive switch (see [Dart Docs (sealed class modifier)](https://dart.dev/language/class-modifiers#sealed)).

### Added

- Added support for date, time and date-time values
- Added support for custom number formats
- Strict typing for cell values (that allow for exhaustive switch statements)

### Fixed

- Issue where seamingly random values are converted to a date iso8601 string, caused by incorrect interpretation of numFmtId=164
- Fixed corrupt excel file when writing large datasets with improvements in shared_strings


## [3.0.0] - 2023-07-30

### Breaking Changes

- Renamed `getColAutoFits()` to `getColumnAutoFits()`, and changed return type to `Map<int, bool>` in `Sheet`
- Renamed `getColWidths()` to `getColumnWidths()`, and changed return type to `Map<int, double>` in `Sheet`
- Renamed `getColAutoFit()` to `getColumnAutoFit()` in `Sheet`
- Renamed `getColWidth()` to `getColumnWidth()` in `Sheet`
- Renamed `setColAutoFit()` to `setColumnAutoFit()` in `Sheet`
- Renamed `setColWidth()` to `setColumnWidth()` in `Sheet`

### Added

- Add setMergedCellStyle() to Sheet, allowing to set style for merged cells
- Add setDefaultRowHeight(), setDefaultColumnWidth() to Sheet
- Add defaultRowHeight and defaultColumnWidth properties to Sheet
- Add getRowHeights(), getRowHeight() and setRowHeight to Sheet
- Add pub topics

### Improved

- Support sharedStrings absolute path
- Loosen up dependency constraints
- Clean up markdown files
- Clean up code

### Fixed

- Fixed many instances of missing/wrong data by comparing strings instead of hashes
- Ignore shared text in 'rPh' element
- Fix findAndReplace() not doing anything

## [2.1.0] - 2023-03-30

### Improved

- Add border functionality

### Fixed

- Fix Header and Footer with special characters
- Fix sheet.merge()

## [2.0.4] - 2023-03-12

### Improved

- Automated Publishing.

## [2.0.3] - 2023-03-12

### Improved

- Readme updated.

## [2.0.2] - 2023-03-12

### Improved

- Fix bug on header and footer.

## [2.0.0-null-safety-4] - 2022-02-15

### Improved

- Fix saving XLXS bug on archive 3.2.0

## [2.0.0-null-safety-3] - 2021-04-29

### Improved

- Forcefully initializing the variables on re-creation

## [2.0.0-null-safety-2] - 2021-04-29

### Improved

- Fix of sharedStringTarget fail to initialize issue

## [2.0.0-null-safety-1] - 2021-04-29

### Improved

- Fix of value not updating in cell

## [2.0.0-null-safety] - 2021-03-28

### Improved

- Null-safety

## [1.1.5] - 2020-08-17

### Improved

- Fixes

## [1.1.4] - 2020-07-23

### Improved

- Improvement in speed of apeending the rows

## [1.1.3] - 2020-07-23

### Improved

- Improvement in speed of apeending the rows

## [1.1.2] - 2020-07-18

### Improved

- Iterating Sheet's Data Object to operate on particular cells

## [1.1.1] - 2020-07-18

### Improved

- Health Improvement

## [1.1.0] - 2020-06-26

### Improved

- Bugs on deleting sheet

## [1.0.9] - 2020-06-06

### Added

- Copy
- Rename
- Delete
- Link Sheets
- Un-Link Sheets
- Font Family
- Font Size
- Italic
- Underline
- Bold

### Improved

- Faster Processing

## [1.0.8] - 2020-05-23

### Removed

- Bugs related to appendRows

## [1.0.7] - 2020-05-21

### Removed

- Bugs related to removal of rows

## [1.0.6] - 2020-05-21

### Added Functionality

- Find and Replace
- Add row / column from Iterables

## [1.0.5] - 2020-05-15

### Removed

- Bugs related to Spanning
- Unwanted removal of rows and columns from spanned cells

## [1.0.4] - 2020-05-10

### Improved

- Analysis related changes
- Vertical Alignment Issue

## [1.0.3] - 2020-05-10

### Added

- Merging of Rows and Columns
- Un-Merging of Rows and Columns
- Font Color
- Background Color
- Setting Default Sheet

## [1.0.2] - 2020-02-18

### Improved

- Minor Bugs

## [1.0.1] - 2020-02-18

### Added

- TextWrapping and (Clip in Google Sheets) / (ShrinkToFit in Microsoft Excel)
- Horizontal and Vertical Alignment
- Update Cell by Cell-Name ("A1")

### Improved

- Health Maintenance

### Fixes

- Minor Bug Fixes

## [1.0.0] - 2020-02-18

- Initial Release
