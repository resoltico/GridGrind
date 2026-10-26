package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;

/** Shared catalog and snapshot helpers for factual workbook table metadata. */
final class ExcelTableCatalogSupport {
  private ExcelTableCatalogSupport() {}

  /** Selects factual table snapshots by workbook-global name in request order. */
  static List<TableEntryReport> selectTablesByName(
      List<TableEntryReport> allTables, List<String> names) {
    Objects.requireNonNull(allTables, "allTables must not be null");
    Objects.requireNonNull(names, "names must not be null");

    List<TableEntryReport> selected = new ArrayList<>();
    for (String name : names) {
      findTableSnapshotByName(allTables, name).ifPresent(selected::add);
    }
    return List.copyOf(selected);
  }

  /** Returns one factual table snapshot by workbook-global name or null when absent. */
  static Optional<TableEntryReport> findTableSnapshotByName(
      List<TableEntryReport> allTables, String name) {
    Objects.requireNonNull(allTables, "allTables must not be null");
    Objects.requireNonNull(name, "name must not be null");

    String expectedName = name.toUpperCase(Locale.ROOT);
    for (TableEntryReport snapshot : allTables) {
      if (snapshot.name().value().toUpperCase(Locale.ROOT).equals(expectedName)) {
        return Optional.of(snapshot);
      }
    }
    return Optional.empty();
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
  static TableEntryReport toSnapshot(String sheetName, XSSFTable table) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(table, "table must not be null");

    CTTable ctTable = table.getCTTable();
    return new TableEntryReport(
        new dev.erst.gridgrind.core.vocabulary.TableLabel(
            Objects.requireNonNullElse(table.getName(), "")),
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName),
        new dev.erst.gridgrind.core.vocabulary.RangeLabel(
            Objects.requireNonNullElse(ctTable.getRef(), "")),
        new TableEntryReport.Structure(
            table.getHeaderRowCount(),
            table.getTotalsRowCount(),
            table.getColumns().stream()
                .map(column -> Objects.requireNonNullElse(column.getName(), ""))
                .toList(),
            Arrays.stream(ctTable.getTableColumns().getTableColumnArray())
                .map(ExcelTableCatalogSupport::toColumnSnapshot)
                .toList()),
        toStyleSnapshot(table),
        new TableEntryReport.Behavior(
            ctTable.isSetAutoFilter(),
            ctTable.getPublished(),
            ctTable.getInsertRow(),
            ctTable.getInsertRowShift()),
        new TableEntryReport.Presentation(
            blankAsOptional(ctTable.getComment()),
            blankAsOptional(ctTable.getHeaderRowCellStyle()),
            blankAsOptional(ctTable.getDataCellStyle()),
            blankAsOptional(ctTable.getTotalsRowCellStyle())));
  }

  /** Converts POI table-style metadata into the factual GridGrind snapshot shape. */
  static TableStyleReport toStyleSnapshot(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    CTTable ctTable = table.getCTTable();
    if (!ctTable.isSetTableStyleInfo()) {
      return new TableStyleReport.None();
    }
    CTTableStyleInfo styleInfo = ctTable.getTableStyleInfo();
    return new TableStyleReport.Named(
        Objects.requireNonNullElse(styleInfo.getName(), ""),
        styleInfo.getShowFirstColumn(),
        styleInfo.getShowLastColumn(),
        styleInfo.getShowRowStripes(),
        styleInfo.getShowColumnStripes());
  }

  private static TableColumnReport toColumnSnapshot(CTTableColumn column) {
    return new TableColumnReport(
        column.getId(),
        Objects.requireNonNullElse(column.getName(), ""),
        Optional.ofNullable(Objects.requireNonNullElse(column.getUniqueName(), "")),
        Optional.ofNullable(Objects.requireNonNullElse(column.getTotalsRowLabel(), "")),
        Optional.ofNullable(
            column.isSetTotalsRowFunction() ? column.getTotalsRowFunction().toString() : ""),
        Optional.ofNullable(
            column.isSetCalculatedColumnFormula()
                ? Objects.requireNonNullElse(
                    column.getCalculatedColumnFormula().getStringValue(), "")
                : ""));
  }

  private static Optional<String> blankAsOptional(String value) {
    if (value == null || value.isBlank()) {
      return Optional.empty();
    }
    return Optional.of(value);
  }
}
