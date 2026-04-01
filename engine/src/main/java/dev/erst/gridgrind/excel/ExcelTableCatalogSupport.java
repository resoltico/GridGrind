package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

/** Shared catalog and snapshot helpers for factual workbook table metadata. */
final class ExcelTableCatalogSupport {
  private ExcelTableCatalogSupport() {}

  /** Selects factual table snapshots by workbook-global name in request order. */
  static List<ExcelTableSnapshot> selectTablesByName(
      List<ExcelTableSnapshot> allTables, List<String> names) {
    Objects.requireNonNull(allTables, "allTables must not be null");
    Objects.requireNonNull(names, "names must not be null");

    List<ExcelTableSnapshot> selected = new ArrayList<>();
    for (String name : names) {
      ExcelTableSnapshot snapshot = findTableSnapshotByName(allTables, name);
      if (snapshot != null) {
        selected.add(snapshot);
      }
    }
    return List.copyOf(selected);
  }

  /** Returns one factual table snapshot by workbook-global name or null when absent. */
  static ExcelTableSnapshot findTableSnapshotByName(
      List<ExcelTableSnapshot> allTables, String name) {
    Objects.requireNonNull(allTables, "allTables must not be null");
    Objects.requireNonNull(name, "name must not be null");

    String expectedName = name.toUpperCase(Locale.ROOT);
    for (ExcelTableSnapshot snapshot : allTables) {
      if (snapshot.name().toUpperCase(Locale.ROOT).equals(expectedName)) {
        return snapshot;
      }
    }
    return null;
  }

  /** Returns the named table from one sheet or throws when the sheet does not contain it. */
  static XSSFTable requiredTableByName(XSSFSheet sheet, String name) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(name, "name must not be null");

    for (XSSFTable table : sheet.getTables()) {
      if (Objects.requireNonNullElse(table.getName(), "").equalsIgnoreCase(name)) {
        return table;
      }
    }
    throw new IllegalArgumentException(
        "table not found on sheet: " + name + "@" + sheet.getSheetName());
  }

  /** Converts one mutable POI table into the factual GridGrind snapshot shape. */
  static ExcelTableSnapshot toSnapshot(String sheetName, XSSFTable table) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(table, "table must not be null");

    CTTable ctTable = table.getCTTable();
    return new ExcelTableSnapshot(
        Objects.requireNonNullElse(table.getName(), ""),
        sheetName,
        Objects.requireNonNullElse(ctTable.getRef(), ""),
        table.getHeaderRowCount(),
        table.getTotalsRowCount(),
        table.getColumns().stream()
            .map(column -> Objects.requireNonNullElse(column.getName(), ""))
            .toList(),
        toStyleSnapshot(table),
        ctTable.isSetAutoFilter());
  }

  /** Converts POI table-style metadata into the factual GridGrind snapshot shape. */
  static ExcelTableStyleSnapshot toStyleSnapshot(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    CTTable ctTable = table.getCTTable();
    if (!ctTable.isSetTableStyleInfo()) {
      return new ExcelTableStyleSnapshot.None();
    }
    CTTableStyleInfo styleInfo = ctTable.getTableStyleInfo();
    return new ExcelTableStyleSnapshot.Named(
        Objects.requireNonNullElse(styleInfo.getName(), ""),
        styleInfo.getShowFirstColumn(),
        styleInfo.getShowLastColumn(),
        styleInfo.getShowRowStripes(),
        styleInfo.getShowColumnStripes());
  }
}
