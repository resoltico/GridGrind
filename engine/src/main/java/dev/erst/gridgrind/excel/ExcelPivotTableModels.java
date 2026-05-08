package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

/** Package-local pivot-table helper models shared across controller support classes. */
enum ResolvedAuthoringSourceKind {
  RANGE,
  NAMED_RANGE,
  TABLE
}

record ResolvedAuthoringSource(
    ResolvedAuthoringSourceKind kind,
    XSSFSheet sheet,
    AreaReference area,
    String description,
    Optional<Name> namedRange,
    Optional<XSSFTable> table) {
  ResolvedAuthoringSource {
    Objects.requireNonNull(kind, "kind must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(area, "area must not be null");
    Objects.requireNonNull(description, "description must not be null");
    Objects.requireNonNull(namedRange, "namedRange must not be null");
    Objects.requireNonNull(table, "table must not be null");
  }

  static ResolvedAuthoringSource range(XSSFSheet sheet, AreaReference area) {
    return new ResolvedAuthoringSource(
        ResolvedAuthoringSourceKind.RANGE,
        sheet,
        area,
        sheet.getSheetName() + "!" + area.formatAsString(),
        Optional.empty(),
        Optional.empty());
  }

  static ResolvedAuthoringSource namedRange(XSSFSheet sheet, AreaReference area, Name namedRange) {
    return new ResolvedAuthoringSource(
        ResolvedAuthoringSourceKind.NAMED_RANGE,
        sheet,
        area,
        "named range " + namedRange.getNameName(),
        Optional.of(namedRange),
        Optional.empty());
  }

  static ResolvedAuthoringSource table(XSSFSheet sheet, AreaReference area, XSSFTable table) {
    return new ResolvedAuthoringSource(
        ResolvedAuthoringSourceKind.TABLE,
        sheet,
        area,
        "table " + table.getName(),
        Optional.empty(),
        Optional.of(table));
  }
}

record ColumnAxisSnapshot(
    List<ExcelPivotTableSnapshot.Field> columnLabels, boolean valuesAxisOnColumns) {
  ColumnAxisSnapshot {
    columnLabels = List.copyOf(columnLabels);
  }
}

record SourceColumn(String name, int relativeIndex) {}

record SourceColumns(List<SourceColumn> columns) {
  SourceColumns {
    columns = List.copyOf(columns);
  }

  int relativeIndex(String name) {
    String expected = name.toUpperCase(Locale.ROOT);
    for (SourceColumn column : columns) {
      if (column.name().toUpperCase(Locale.ROOT).equals(expected)) {
        return column.relativeIndex();
      }
    }
    throw new IllegalArgumentException("pivot source column not found: " + name);
  }
}

record PivotHandle(
    int sheetIndex, int ordinalOnSheet, String sheetName, XSSFSheet sheet, XSSFPivotTable table) {}

record PivotLocation(String topLeftAddress, String locationRange) {}
