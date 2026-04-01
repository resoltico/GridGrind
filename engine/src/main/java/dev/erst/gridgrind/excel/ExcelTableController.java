package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;

/** Reads, writes, and analyzes workbook tables. */
final class ExcelTableController {
  private final ExcelAutofilterController autofilterController;

  ExcelTableController() {
    this(new ExcelAutofilterController());
  }

  ExcelTableController(ExcelAutofilterController autofilterController) {
    this.autofilterController =
        Objects.requireNonNull(autofilterController, "autofilterController must not be null");
  }

  /** Creates or replaces one workbook-global table definition. */
  void setTable(ExcelWorkbook workbook, ExcelTableDefinition definition) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    ExcelRange targetRange = ExcelRange.parse(definition.range());
    ExcelTableStructureSupport.requireSupportedTableShape(targetRange, definition.showTotalsRow());
    XSSFSheet sheet = requiredSheet(workbook, definition.sheetName());
    List<String> headerNames = ExcelTableStructureSupport.headerNames(sheet, targetRange);
    requireHeaders(headerNames);

    TableHandle existingByName = tableByName(workbook, definition.name());
    if (existingByName != null && !existingByName.sheetName().equals(definition.sheetName())) {
      throw new IllegalArgumentException(
          "table name already exists on a different sheet: " + definition.name());
    }
    requireNamedRangeNameAvailable(workbook, definition.name());
    requireNoOverlappingOtherTables(workbook, definition, targetRange, existingByName);
    validateStyle(workbook, definition.style());

    if (existingByName != null) {
      existingByName.sheet().removeTable(existingByName.table());
    }

    XSSFTable table =
        sheet.createTable(
            new AreaReference(
                ExcelSheetStructureSupport.formatRange(targetRange), SpreadsheetVersion.EXCEL2007));
    CTTable ctTable = table.getCTTable();
    ctTable.setHeaderRowCount(1);
    ctTable.setTotalsRowCount(definition.showTotalsRow() ? 1 : 0);
    ctTable.setTotalsRowShown(definition.showTotalsRow());
    table.setDisplayName(definition.name());
    table.setName(definition.name());
    ExcelTableStructureSupport.applyAutofilter(table, targetRange, definition.showTotalsRow());
    ExcelTableStructureSupport.applyStyle(table, definition.style());
    table.updateHeaders();
    clearOverlappingSheetAutofilter(sheet, targetRange);
  }

  /** Deletes one existing table by workbook-global name and expected sheet name. */
  void deleteTable(ExcelWorkbook workbook, String name, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String validatedName = ExcelTableDefinition.validateName(name);
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }

    TableHandle tableHandle = tableByName(workbook, validatedName);
    if (tableHandle == null || !tableHandle.sheetName().equals(sheetName)) {
      throw new IllegalArgumentException(
          "table not found on expected sheet: " + validatedName + "@" + sheetName);
    }
    tableHandle.sheet().removeTable(tableHandle.table());
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  List<ExcelTableSnapshot> tables(ExcelWorkbook workbook, ExcelTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelTableSnapshot> allTables =
        allTables(workbook).stream()
            .map(
                tableHandle ->
                    ExcelTableCatalogSupport.toSnapshot(
                        tableHandle.sheetName(), tableHandle.table()))
            .toList();
    return switch (selection) {
      case ExcelTableSelection.All _ -> allTables;
      case ExcelTableSelection.ByNames byNames ->
          ExcelTableCatalogSupport.selectTablesByName(allTables, byNames.names());
    };
  }

  /** Returns table-owned autofilter metadata present on one sheet. */
  List<ExcelAutofilterSnapshot> tableOwnedAutofilters(ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");

    List<ExcelAutofilterSnapshot> autofilters = new ArrayList<>();
    for (TableHandle tableHandle : allTables(workbook)) {
      if (!tableHandle.sheetName().equals(sheetName)) {
        continue;
      }
      if (!tableHandle.table().getCTTable().isSetAutoFilter()) {
        continue;
      }
      String rawRange =
          Objects.requireNonNullElse(tableHandle.table().getCTTable().getAutoFilter().getRef(), "");
      autofilters.add(
          new ExcelAutofilterSnapshot.TableOwned(rawRange, tableHandle.table().getName()));
    }
    return List.copyOf(autofilters);
  }

  /** Returns the number of table-owned autofilters present on one sheet. */
  int tableAutofilterCount(ExcelWorkbook workbook, String sheetName) {
    return tableOwnedAutofilters(workbook, sheetName).size();
  }

  /** Returns derived table-health findings for the selected tables. */
  List<WorkbookAnalysis.AnalysisFinding> tableHealthFindings(
      ExcelWorkbook workbook, ExcelTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelTableSnapshot> selectedTables = tables(workbook, selection);
    List<ExcelTableSnapshot> allTables = tables(workbook, new ExcelTableSelection.All());
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (ExcelTableSnapshot table : selectedTables) {
      findings.addAll(ExcelTableAnalysisSupport.tableFindings(workbook, table));
      findings.addAll(ExcelTableAnalysisSupport.overlapFindings(table, allTables));
    }
    return List.copyOf(findings.stream().distinct().toList());
  }

  /** Returns derived autofilter-health findings for table-owned autofilters on one sheet. */
  List<WorkbookAnalysis.AnalysisFinding> tableAutofilterHealthFindings(
      ExcelWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");

    XSSFSheet sheet = requiredSheet(workbook, sheetName);
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (ExcelTableSnapshot table : tables(workbook, new ExcelTableSelection.All())) {
      if (!table.sheetName().equals(sheetName) || !table.hasAutofilter()) {
        continue;
      }
      findings.addAll(ExcelTableAnalysisSupport.tableAutofilterFindings(sheet, table));
    }
    return List.copyOf(findings);
  }

  private void clearOverlappingSheetAutofilter(XSSFSheet sheet, ExcelRange tableRange) {
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return;
    }
    ExcelRange sheetAutofilterRange =
        ExcelSheetStructureSupport.parseRangeOrNull(
            Objects.requireNonNullElse(sheet.getCTWorksheet().getAutoFilter().getRef(), ""));
    if (sheetAutofilterRange != null
        && ExcelSheetStructureSupport.intersects(sheetAutofilterRange, tableRange)) {
      autofilterController.clearSheetAutofilter(sheet);
    }
  }

  private void requireNoOverlappingOtherTables(
      ExcelWorkbook workbook,
      ExcelTableDefinition definition,
      ExcelRange targetRange,
      TableHandle existingByName) {
    List<String> overlaps = new ArrayList<>();
    for (TableHandle tableHandle : allTables(workbook)) {
      if (!tableHandle.sheetName().equals(definition.sheetName())) {
        continue;
      }
      if (existingByName != null && tableHandle.name().equalsIgnoreCase(existingByName.name())) {
        continue;
      }
      ExcelRange existingRange =
          ExcelSheetStructureSupport.parseRangeOrNull(
              Objects.requireNonNullElse(tableHandle.table().getCTTable().getRef(), ""));
      if (existingRange != null
          && ExcelSheetStructureSupport.intersects(existingRange, targetRange)) {
        overlaps.add(tableHandle.name() + "@" + tableHandle.table().getCTTable().getRef());
      }
    }
    if (!overlaps.isEmpty()) {
      throw new IllegalArgumentException(
          "table range must not overlap an existing table: " + String.join(", ", overlaps));
    }
  }

  private void requireNamedRangeNameAvailable(ExcelWorkbook workbook, String tableName) {
    for (Name name : workbook.xssfWorkbook().getAllNames()) {
      if (Objects.requireNonNullElse(name.getNameName(), "").equalsIgnoreCase(tableName)) {
        throw new IllegalArgumentException(
            "table name must not conflict with an existing defined name: " + tableName);
      }
    }
  }

  private void requireHeaders(List<String> headerNames) {
    Set<String> seen = new LinkedHashSet<>();
    for (String headerName : headerNames) {
      if (headerName.isBlank()) {
        throw new IllegalArgumentException("table header cells must not be blank");
      }
      String key = headerName.toUpperCase(Locale.ROOT);
      if (!seen.add(key)) {
        throw new IllegalArgumentException(
            "table header cells must be unique (case-insensitive): " + headerName);
      }
    }
  }

  private void validateStyle(ExcelWorkbook workbook, ExcelTableStyle style) {
    switch (style) {
      case ExcelTableStyle.None _ -> {}
      case ExcelTableStyle.Named named -> {
        if (workbook.xssfWorkbook().getStylesSource().getTableStyle(named.name()) == null) {
          throw new IllegalArgumentException("unknown table style: " + named.name());
        }
      }
    }
  }

  private List<TableHandle> allTables(ExcelWorkbook workbook) {
    List<TableHandle> tables = new ArrayList<>();
    for (String sheetName : workbook.sheetNames()) {
      XSSFSheet sheet = requiredSheet(workbook, sheetName);
      for (XSSFTable table : sheet.getTables()) {
        tables.add(
            new TableHandle(
                Objects.requireNonNullElse(table.getName(), ""), sheetName, sheet, table));
      }
    }
    return List.copyOf(tables);
  }

  private TableHandle tableByName(ExcelWorkbook workbook, String name) {
    String key = name.toUpperCase(Locale.ROOT);
    for (TableHandle tableHandle : allTables(workbook)) {
      if (tableHandle.name().toUpperCase(Locale.ROOT).equals(key)) {
        return tableHandle;
      }
    }
    return null;
  }

  private static XSSFSheet requiredSheet(ExcelWorkbook workbook, String sheetName) {
    XSSFSheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }

  private record TableHandle(String name, String sheetName, XSSFSheet sheet, XSSFTable table) {
    private TableHandle {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(sheetName, "sheetName must not be null");
      Objects.requireNonNull(sheet, "sheet must not be null");
      Objects.requireNonNull(table, "table must not be null");
    }
  }
}
