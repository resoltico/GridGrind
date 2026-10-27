package dev.erst.gridgrind.engine.tables;

import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import dev.erst.gridgrind.engine.backend.xssf.SheetNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.FormulaWrites;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction;

/** Owns table mutation, factual catalogs, header/formula synchronization and derived health. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class TableAdapter {
  private final AutofilterAdapter autofilterAdapter;

  public TableAdapter() {
    this(new AutofilterAdapter());
  }

  public TableAdapter(AutofilterAdapter autofilterAdapter) {
    this.autofilterAdapter =
        Objects.requireNonNull(autofilterAdapter, "autofilterAdapter must not be null");
  }

  /** Creates or replaces one workbook-global table definition. */
  public void setTable(XssfWorkbook workbook, ResolvedTable definition) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(definition, "definition must not be null");

    CellBounds targetRange = CellBounds.parse(definition.range().value());
    requireSupportedTableShape(targetRange, definition.showTotalsRow());
    XSSFSheet sheet = requiredSheet(workbook, definition.sheetName().value());
    List<String> headerNames = SheetStructure.headerNames(sheet, targetRange);
    requireHeaders(headerNames);

    Optional<TableHandle> existingByName = tableByName(workbook, definition.name().value());
    if (existingByName.isPresent()
        && !existingByName.orElseThrow().sheetName().equals(definition.sheetName().value())) {
      throw new IllegalArgumentException(
          "table name already exists on a different sheet: " + definition.name().value());
    }
    requireNamedRangeNameAvailable(workbook, definition.name().value());
    requireNoOverlappingOtherTables(workbook, definition, targetRange, existingByName);
    validateStyle(workbook, definition.style());

    XSSFTable table =
        existingByName.isEmpty()
            ? sheet.createTable(
                new AreaReference(
                    SheetStructure.formatRange(targetRange), SpreadsheetVersion.EXCEL2007))
            : existingByName.orElseThrow().table();
    if (existingByName.isPresent()) {
      table.setArea(
          new AreaReference(SheetStructure.formatRange(targetRange), SpreadsheetVersion.EXCEL2007));
    }
    CTTable ctTable = table.getCTTable();
    ctTable.setHeaderRowCount(1);
    ctTable.setTotalsRowCount(definition.showTotalsRow() ? 1 : 0);
    ctTable.setTotalsRowShown(definition.showTotalsRow());
    table.setDisplayName(definition.name().value());
    table.setName(definition.name().value());
    applyStyle(table, definition.style());
    table.updateHeaders();
    applyTableMetadata(ctTable, definition);
    applyColumnMetadata(table, definition.columns());
    if (definition.hasAutofilter()) {
      applyAutofilter(table, targetRange, definition.showTotalsRow());
    } else if (ctTable.isSetAutoFilter()) {
      ctTable.unsetAutoFilter();
    }
    clearOverlappingSheetAutofilter(sheet, targetRange);
  }

  /** Deletes one existing table by workbook-global name and expected sheet name. */
  public void deleteTable(XssfWorkbook workbook, String name, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String validatedName = new TableName(name).value();
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }

    Optional<TableHandle> tableHandle = tableByName(workbook, validatedName);
    if (tableHandle.isEmpty() || !tableHandle.orElseThrow().sheetName().equals(sheetName)) {
      throw new IllegalArgumentException(
          "table not found on expected sheet: " + validatedName + "@" + sheetName);
    }
    tableHandle.orElseThrow().sheet().removeTable(tableHandle.orElseThrow().table());
  }

  /** Returns factual table metadata selected by workbook-global table name or all tables. */
  public List<TableEntryReport> tables(XssfWorkbook workbook, TableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<TableEntryReport> allTables =
        allTables(workbook).stream()
            .map(tableHandle -> toSnapshot(tableHandle.sheetName(), tableHandle.table()))
            .toList();
    return switch (selection) {
      case TableSelector.All _ -> allTables;
      case TableSelector.ByNames byNames ->
          selectTablesByName(
              allTables,
              byNames.names().stream()
                  .map(dev.erst.gridgrind.core.vocabulary.TableName::value)
                  .toList());
      case TableSelector.ByName byName ->
          selectTablesByName(allTables, List.of(byName.name().value()));
      case TableSelector.ByNameOnSheet scoped ->
          selectTablesByName(allTables, List.of(scoped.name().value())).stream()
              .filter(
                  table -> table.sheetName().value().equalsIgnoreCase(scoped.sheetName().value()))
              .toList();
    };
  }

  /** Returns table-owned autofilter metadata present on one sheet. */
  public List<AutofilterEntryReport> tableOwnedAutofilters(
      XssfWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");

    List<AutofilterEntryReport> autofilters = new ArrayList<>();
    for (TableHandle tableHandle : allTables(workbook)) {
      if (!tableHandle.sheetName().equals(sheetName)) {
        continue;
      }
      if (!tableHandle.table().getCTTable().isSetAutoFilter()) {
        continue;
      }
      var autoFilter = tableHandle.table().getCTTable().getAutoFilter();
      autofilters.add(
          new AutofilterEntryReport.TableOwned(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel(
                  Objects.requireNonNullElse(autoFilter.getRef(), "")),
              tableHandle.table().getName(),
              AutofilterAdapter.filterColumns(workbook.xssfWorkbook(), autoFilter),
              AutofilterAdapter.sortState(workbook.xssfWorkbook(), autoFilter)));
    }
    return List.copyOf(autofilters);
  }

  /** Returns the number of table-owned autofilters present on one sheet. */
  public int tableAutofilterCount(XssfWorkbook workbook, String sheetName) {
    return tableOwnedAutofilters(workbook, sheetName).size();
  }

  /** Returns derived table-health findings for the selected tables. */
  public List<AnalysisFindingReport> tableHealthFindings(
      XssfWorkbook workbook, TableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<TableEntryReport> selectedTables = tables(workbook, selection);
    List<TableEntryReport> allTables = tables(workbook, new TableSelector.All());
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (TableEntryReport table : selectedTables) {
      findings.addAll(tableFindings(workbook, table));
      findings.addAll(overlapFindings(table, allTables));
    }
    return List.copyOf(findings.stream().distinct().toList());
  }

  /** Returns derived autofilter-health findings for table-owned autofilters on one sheet. */
  public List<AnalysisFindingReport> tableAutofilterHealthFindings(
      XssfWorkbook workbook, String sheetName) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");

    XSSFSheet sheet = requiredSheet(workbook, sheetName);
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (TableEntryReport table : tables(workbook, new TableSelector.All())) {
      if (!table.sheetName().value().equals(sheetName) || !table.behavior().hasAutofilter()) {
        continue;
      }
      findings.addAll(tableAutofilterFindings(sheet, table));
    }
    return List.copyOf(findings);
  }

  private void clearOverlappingSheetAutofilter(XSSFSheet sheet, CellBounds tableRange) {
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return;
    }
    Optional<CellBounds> sheetAutofilterRange =
        SheetStructure.parseOptionalRange(
            Objects.requireNonNullElse(sheet.getCTWorksheet().getAutoFilter().getRef(), ""));
    if (sheetAutofilterRange.isPresent()
        && SheetStructure.intersects(sheetAutofilterRange.orElseThrow(), tableRange)) {
      autofilterAdapter.clearSheetAutofilter(sheet);
    }
  }

  private void requireNoOverlappingOtherTables(
      XssfWorkbook workbook,
      ResolvedTable definition,
      CellBounds targetRange,
      Optional<TableHandle> existingByName) {
    List<String> overlaps = new ArrayList<>();
    for (TableHandle tableHandle : allTables(workbook)) {
      if (!tableHandle.sheetName().equals(definition.sheetName().value())) {
        continue;
      }
      if (existingByName.isPresent()
          && tableHandle.name().equalsIgnoreCase(existingByName.orElseThrow().name())) {
        continue;
      }
      Optional<CellBounds> existingRange =
          SheetStructure.parseOptionalRange(
              Objects.requireNonNullElse(tableHandle.table().getCTTable().getRef(), ""));
      if (existingRange.isPresent()
          && SheetStructure.intersects(existingRange.orElseThrow(), targetRange)) {
        overlaps.add(tableHandle.name() + "@" + tableHandle.table().getCTTable().getRef());
      }
    }
    if (!overlaps.isEmpty()) {
      throw new IllegalArgumentException(
          "table range must not overlap an existing table: " + String.join(", ", overlaps));
    }
  }

  private void requireNamedRangeNameAvailable(XssfWorkbook workbook, String tableName) {
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

  private void validateStyle(XssfWorkbook workbook, TableStyleSpec style) {
    switch (style) {
      case TableStyleSpec.None _ -> {}
      case TableStyleSpec.Named named -> {
        if (workbook.xssfWorkbook().getStylesSource().getTableStyle(named.name()) == null) {
          throw new IllegalArgumentException("unknown table style: " + named.name());
        }
      }
    }
  }

  private static void applyTableMetadata(CTTable ctTable, ResolvedTable definition) {
    setOrUnset(
        definition.comment(),
        ctTable::setComment,
        () -> {
          if (ctTable.isSetComment()) {
            ctTable.unsetComment();
          }
        });
    ctTable.setPublished(definition.published());
    ctTable.setInsertRow(definition.insertRow());
    ctTable.setInsertRowShift(definition.insertRowShift());
    setOrUnset(
        definition.headerRowCellStyle(),
        ctTable::setHeaderRowCellStyle,
        () -> {
          if (ctTable.isSetHeaderRowCellStyle()) {
            ctTable.unsetHeaderRowCellStyle();
          }
        });
    setOrUnset(
        definition.dataCellStyle(),
        ctTable::setDataCellStyle,
        () -> {
          if (ctTable.isSetDataCellStyle()) {
            ctTable.unsetDataCellStyle();
          }
        });
    setOrUnset(
        definition.totalsRowCellStyle(),
        ctTable::setTotalsRowCellStyle,
        () -> {
          if (ctTable.isSetTotalsRowCellStyle()) {
            ctTable.unsetTotalsRowCellStyle();
          }
        });
  }

  private static void applyColumnMetadata(XSSFTable table, List<TableColumnSpec> columns) {
    for (TableColumnSpec definition : columns) {
      CTTableColumn column = requiredColumn(table, definition.columnIndex());
      applyColumnUniqueName(column, definition.uniqueName());
      applyTotalsRowLabel(column, definition.totalsRowLabel());
      applyTotalsRowFunction(column, definition.totalsRowFunction());
      applyCalculatedColumnFormula(column, definition.calculatedColumnFormula());
    }
  }

  private static CTTableColumn requiredColumn(XSSFTable table, int columnIndex) {
    if (columnIndex >= table.getColumns().size()) {
      throw new IllegalArgumentException(
          "table columnIndex is outside the table range: " + columnIndex);
    }
    return table.getCTTable().getTableColumns().getTableColumnArray(columnIndex);
  }

  private static void applyColumnUniqueName(CTTableColumn column, String uniqueName) {
    setOrUnset(
        uniqueName,
        column::setUniqueName,
        () -> {
          if (column.isSetUniqueName()) {
            column.unsetUniqueName();
          }
        });
  }

  private static void applyTotalsRowLabel(CTTableColumn column, String totalsRowLabel) {
    setOrUnset(
        totalsRowLabel,
        column::setTotalsRowLabel,
        () -> {
          if (column.isSetTotalsRowLabel()) {
            column.unsetTotalsRowLabel();
          }
        });
  }

  private static void applyTotalsRowFunction(CTTableColumn column, String totalsRowFunction) {
    if (totalsRowFunction.isBlank()) {
      if (column.isSetTotalsRowFunction()) {
        column.unsetTotalsRowFunction();
      }
      return;
    }
    STTotalsRowFunction.Enum resolved = STTotalsRowFunction.Enum.forString(totalsRowFunction);
    if (resolved == null) {
      throw new IllegalArgumentException(
          "unsupported table totalsRowFunction: " + totalsRowFunction);
    }
    column.setTotalsRowFunction(resolved);
  }

  private static void applyCalculatedColumnFormula(CTTableColumn column, String formulaText) {
    if (formulaText.isBlank()) {
      if (column.isSetCalculatedColumnFormula()) {
        column.unsetCalculatedColumnFormula();
      }
      return;
    }
    var formula =
        column.isSetCalculatedColumnFormula()
            ? column.getCalculatedColumnFormula()
            : column.addNewCalculatedColumnFormula();
    formula.setStringValue(formulaText);
  }

  private static void setOrUnset(
      String value, java.util.function.Consumer<String> setter, Runnable unsetter) {
    if (value.isBlank()) {
      unsetter.run();
    } else {
      setter.accept(value);
    }
  }

  private List<TableHandle> allTables(XssfWorkbook workbook) {
    List<TableHandle> tables = new ArrayList<>();
    for (String sheetName : workbook.sheets().sheetNames()) {
      XSSFSheet sheet = requiredSheet(workbook, sheetName);
      for (XSSFTable table : sheet.getTables()) {
        tables.add(
            new TableHandle(
                Objects.requireNonNullElse(table.getName(), ""), sheetName, sheet, table));
      }
    }
    return List.copyOf(tables);
  }

  private Optional<TableHandle> tableByName(XssfWorkbook workbook, String name) {
    String key = name.toUpperCase(Locale.ROOT);
    for (TableHandle tableHandle : allTables(workbook)) {
      if (tableHandle.name().toUpperCase(Locale.ROOT).equals(key)) {
        return Optional.of(tableHandle);
      }
    }
    return Optional.empty();
  }

  private static XSSFSheet requiredSheet(XssfWorkbook workbook, String sheetName) {
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

  /** Selects factual table snapshots by workbook-global name in request order. */
  public static List<TableEntryReport> selectTablesByName(
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
  public static Optional<TableEntryReport> findTableSnapshotByName(
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
  public static XSSFTable requiredTableByName(XSSFSheet sheet, String name) {
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
  public static TableEntryReport toSnapshot(String sheetName, XSSFTable table) {
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
                .map(TableAdapter::toColumnSnapshot)
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
  public static TableStyleReport toStyleSnapshot(XSSFTable table) {
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

  /** Returns derived health findings for one factual table snapshot. */
  public static List<AnalysisFindingReport> tableFindings(
      XssfWorkbook workbook, TableEntryReport table) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(table, "table must not be null");

    List<AnalysisFindingReport> findings = new ArrayList<>();
    Optional<CellBounds> parsedRange = SheetStructure.parseOptionalRange(table.range().value());
    AnalysisLocationReport location = tableLocation(table, parsedRange);
    if (parsedRange.isEmpty()) {
      return brokenReferenceFindings(table, location);
    }

    CellBounds range = parsedRange.orElseThrow();
    addRowCountFinding(findings, table, range, location);
    addHeaderFindings(findings, table, location);
    addStyleFinding(findings, workbook, table, location);
    return List.copyOf(findings);
  }

  private static AnalysisLocationReport tableLocation(
      TableEntryReport table, Optional<CellBounds> parsedRange) {
    return parsedRange.isEmpty()
        ? new AnalysisLocationReport.Sheet(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(table.sheetName().value()))
        : new AnalysisLocationReport.Range(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(table.sheetName().value()),
            SheetStructure.formatRange(parsedRange.orElseThrow()));
  }

  private static List<AnalysisFindingReport> brokenReferenceFindings(
      TableEntryReport table, AnalysisLocationReport location) {
    return List.of(
        new AnalysisFindingReport(
            AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
            AnalysisSeverity.ERROR,
            "Table range is invalid",
            "Table range could not be parsed from workbook metadata.",
            location,
            List.of(table.name().value(), table.range().value())));
  }

  private static void addRowCountFinding(
      List<AnalysisFindingReport> findings,
      TableEntryReport table,
      CellBounds range,
      AnalysisLocationReport location) {
    int minimumRows = table.structure().headerRowCount() + table.structure().totalsRowCount() + 1;
    if (table.structure().headerRowCount() < 1 || range.rowCount() < minimumRows) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
              AnalysisSeverity.ERROR,
              "Table range does not match table row counts",
              "Table metadata requires at least "
                  + minimumRows
                  + " rows but the stored range contains only "
                  + range.rowCount()
                  + ".",
              location,
              List.of(table.range().value())));
    }
  }

  private static void addHeaderFindings(
      List<AnalysisFindingReport> findings,
      TableEntryReport table,
      AnalysisLocationReport location) {
    List<String> duplicateHeaders = new ArrayList<>();
    List<String> blankHeaders = new ArrayList<>();
    Set<String> seen = new LinkedHashSet<>();
    for (String header : table.structure().columnNames()) {
      if (header.isBlank()) {
        blankHeaders.add(header);
        continue;
      }
      if (!seen.add(header.toUpperCase(Locale.ROOT))) {
        duplicateHeaders.add(header);
      }
    }
    if (!blankHeaders.isEmpty()) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.TABLE_BLANK_HEADER,
              AnalysisSeverity.ERROR,
              "Table contains blank header cells",
              "Table column headers must be nonblank.",
              location,
              List.of(table.name().value(), table.range().value())));
    }
    if (!duplicateHeaders.isEmpty()) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.TABLE_DUPLICATE_HEADER,
              AnalysisSeverity.ERROR,
              "Table contains duplicate header cells",
              "Table column headers must be unique (case-insensitive).",
              location,
              List.copyOf(duplicateHeaders)));
    }
  }

  private static void addStyleFinding(
      List<AnalysisFindingReport> findings,
      XssfWorkbook workbook,
      TableEntryReport table,
      AnalysisLocationReport location) {
    switch (table.style()) {
      case TableStyleReport.None _ -> {}
      case TableStyleReport.Named named -> {
        if (named.name().isBlank()
            || workbook.xssfWorkbook().getStylesSource().getTableStyle(named.name()) == null) {
          findings.add(
              new AnalysisFindingReport(
                  AnalysisFindingCode.TABLE_STYLE_MISMATCH,
                  AnalysisSeverity.WARNING,
                  "Table style does not resolve",
                  "Table style metadata refers to a style name that the workbook does not define.",
                  location,
                  List.of(named.name())));
        }
      }
    }
  }

  /** Returns derived overlap findings for one factual table snapshot against a peer set. */
  public static List<AnalysisFindingReport> overlapFindings(
      TableEntryReport table, List<TableEntryReport> allTables) {
    Objects.requireNonNull(table, "table must not be null");
    Objects.requireNonNull(allTables, "allTables must not be null");

    Optional<CellBounds> tableRange = SheetStructure.parseOptionalRange(table.range().value());
    if (tableRange.isEmpty()) {
      return List.of();
    }
    List<String> overlaps = new ArrayList<>();
    for (TableEntryReport other : allTables) {
      if (other.equals(table) || !other.sheetName().value().equals(table.sheetName().value())) {
        continue;
      }
      Optional<CellBounds> otherRange = SheetStructure.parseOptionalRange(other.range().value());
      if (otherRange.isPresent()
          && SheetStructure.intersects(tableRange.orElseThrow(), otherRange.orElseThrow())) {
        overlaps.add(other.name().value() + "@" + other.range().value());
      }
    }
    if (overlaps.isEmpty()) {
      return List.of();
    }
    List<String> evidence = new ArrayList<>();
    evidence.add(table.name().value() + "@" + table.range().value());
    evidence.addAll(overlaps);
    return List.of(
        new AnalysisFindingReport(
            AnalysisFindingCode.TABLE_OVERLAPPING_RANGE,
            AnalysisSeverity.ERROR,
            "Table range overlaps another table",
            "Table metadata overlaps one or more other tables on the same sheet.",
            new AnalysisLocationReport.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(table.sheetName().value()),
                table.range().value()),
            List.copyOf(evidence)));
  }

  /** Returns derived autofilter findings for one factual table snapshot on one sheet. */
  public static List<AnalysisFindingReport> tableAutofilterFindings(
      XSSFSheet sheet, TableEntryReport table) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(table, "table must not be null");

    Optional<CellBounds> tableRange = SheetStructure.parseOptionalRange(table.range().value());
    AnalysisLocationReport location =
        tableRange.isEmpty()
            ? new AnalysisLocationReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(table.sheetName().value()))
            : new AnalysisLocationReport.Range(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel(table.sheetName().value()),
                expectedAutofilterRangeText(table));

    XSSFTable xssfTable = requiredTableByName(sheet, table.name().value());
    String rawFilterRange =
        Objects.requireNonNullElse(xssfTable.getCTTable().getAutoFilter().getRef(), "");
    Optional<CellBounds> parsedFilterRange = SheetStructure.parseOptionalRange(rawFilterRange);
    if (parsedFilterRange.isEmpty()) {
      return List.of(
          new AnalysisFindingReport(
              AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
              AnalysisSeverity.ERROR,
              "Table autofilter range is invalid",
              "Table-owned autofilter range could not be parsed from workbook metadata.",
              new AnalysisLocationReport.Sheet(
                  new dev.erst.gridgrind.core.vocabulary.SheetLabel(table.sheetName().value())),
              List.of(table.name().value(), rawFilterRange)));
    }

    List<AnalysisFindingReport> findings = new ArrayList<>();
    if (SheetStructure.headerRowMissing(sheet, parsedFilterRange.orElseThrow())) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
              AnalysisSeverity.WARNING,
              "Table autofilter header row is blank",
              "Table-owned autofilter range does not contain a nonblank header row.",
              location,
              List.of(rawFilterRange, table.name().value())));
    }
    String expectedRange = expectedAutofilterRangeText(table);
    if (!SheetStructure.formatRange(parsedFilterRange.orElseThrow()).equals(expectedRange)) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
              AnalysisSeverity.WARNING,
              "Table autofilter does not match table range",
              "Table-owned autofilter range must match the table range excluding any totals row.",
              location,
              List.of(table.name().value(), rawFilterRange, expectedRange)));
    }
    return List.copyOf(findings);
  }

  /** Validates that one table range can hold its header row, data rows, and optional totals row. */
  public static void requireSupportedTableShape(CellBounds range, boolean showTotalsRow) {
    int minimumRows = showTotalsRow ? 3 : 2;
    if (range.rowCount() < minimumRows) {
      throw new IllegalArgumentException(
          "table range must contain at least "
              + minimumRows
              + " rows to cover a header row, one data row, and any totals row");
    }
  }

  /** Applies one GridGrind table-owned autofilter definition to a mutable XSSF table. */
  public static void applyAutofilter(XSSFTable table, CellBounds range, boolean showTotalsRow) {
    CTAutoFilter autoFilter =
        table.getCTTable().isSetAutoFilter()
            ? table.getCTTable().getAutoFilter()
            : table.getCTTable().addNewAutoFilter();
    int lastFilterRow = showTotalsRow ? range.lastRow() - 1 : range.lastRow();
    autoFilter.setRef(
        new CellRangeAddress(
                range.firstRow(), lastFilterRow, range.firstColumn(), range.lastColumn())
            .formatAsString());
  }

  /** Applies one GridGrind table-style definition to a mutable XSSF table. */
  public static void applyStyle(XSSFTable table, TableStyleSpec style) {
    Objects.requireNonNull(table, "table must not be null");
    Objects.requireNonNull(style, "style must not be null");

    CTTable ctTable = table.getCTTable();
    switch (style) {
      case TableStyleSpec.None _ -> {
        if (ctTable.isSetTableStyleInfo()) {
          ctTable.unsetTableStyleInfo();
        }
      }
      case TableStyleSpec.Named named -> {
        CTTableStyleInfo styleInfo =
            ctTable.isSetTableStyleInfo()
                ? ctTable.getTableStyleInfo()
                : ctTable.addNewTableStyleInfo();
        styleInfo.setName(named.name());
        styleInfo.setShowFirstColumn(named.showFirstColumn());
        styleInfo.setShowLastColumn(named.showLastColumn());
        styleInfo.setShowRowStripes(named.showRowStripes());
        styleInfo.setShowColumnStripes(named.showColumnStripes());
      }
    }
  }

  /** Returns the table-owned autofilter range text implied by one factual table snapshot. */
  public static String expectedAutofilterRangeText(TableEntryReport table) {
    Objects.requireNonNull(table, "table must not be null");

    Optional<CellBounds> range = SheetStructure.parseOptionalRange(table.range().value());
    if (range.isEmpty()) {
      return table.range().value();
    }
    CellBounds parsedRange = range.orElseThrow();
    int lastFilterRow =
        parsedRange.lastRow()
            - Math.min(table.structure().totalsRowCount(), parsedRange.rowCount() - 1);
    return new CellRangeAddress(
            parsedRange.firstRow(),
            lastFilterRow,
            parsedRange.firstColumn(),
            parsedRange.lastColumn())
        .formatAsString();
  }

  /**
   * Clears materialized structured-reference body formulas so POI rename/update logic only sees
   * parseable ordinary cell formulas.
   */
  public static void canonicalizeWorkbook(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      canonicalizeSheet(workbook.getSheetAt(sheetIndex));
    }
  }

  /** Canonicalizes one sheet's table body cells using the authored table metadata as truth. */
  public static void canonicalizeSheet(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    for (XSSFTable table : sheet.getTables()) {
      canonicalizeTable(sheet, table);
    }
  }

  private static void canonicalizeTable(XSSFSheet sheet, XSSFTable table) {
    Optional<CellBounds> range =
        SheetStructure.parseOptionalRange(
            Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
    if (range.isEmpty()) {
      return;
    }
    CellBounds parsedRange = range.orElseThrow();
    int firstBodyRow = parsedRange.firstRow() + 1;
    int lastBodyRow = parsedRange.lastRow() - totalsRowCount(table.getCTTable());
    if (firstBodyRow > lastBodyRow) {
      return;
    }

    CTTableColumn[] columns = table.getCTTable().getTableColumns().getTableColumnArray();
    for (int columnOffset = 0; columnOffset < columns.length; columnOffset++) {
      Optional<String> calculatedFormula = calculatedColumnFormula(columns[columnOffset]);
      if (calculatedFormula.isEmpty()) {
        continue;
      }
      int columnIndex = parsedRange.firstColumn() + columnOffset;
      for (int rowIndex = firstBodyRow; rowIndex <= lastBodyRow; rowIndex++) {
        clearMaterializedCalculatedFormulaCell(
            sheet, rowIndex, columnIndex, calculatedFormula.orElseThrow());
      }
    }
  }

  private static int totalsRowCount(CTTable table) {
    return table.getTotalsRowCount() > 0 ? Math.toIntExact(table.getTotalsRowCount()) : 0;
  }

  public static Optional<String> calculatedColumnFormula(CTTableColumn column) {
    if (!column.isSetCalculatedColumnFormula()) {
      return Optional.empty();
    }
    String formula =
        Objects.requireNonNullElse(column.getCalculatedColumnFormula().getStringValue(), "");
    return formula.isBlank() ? Optional.empty() : Optional.of(formula);
  }

  private static void clearMaterializedCalculatedFormulaCell(
      XSSFSheet sheet, int rowIndex, int columnIndex, String calculatedFormula) {
    Row row = sheet.getRow(rowIndex);
    if (row == null) {
      return;
    }
    Cell cell = row.getCell(columnIndex);
    if (cell == null || cell.getCellType() != CellType.FORMULA) {
      return;
    }
    if (!calculatedFormula.equals(cell.getCellFormula())) {
      return;
    }
    if (mustPreserveCellShell(cell)) {
      cell.setBlank();
      return;
    }
    row.removeCell(cell);
  }

  public static boolean mustPreserveCellShell(Cell cell) {
    return cell.getCellComment() != null
        || cell.getHyperlink() != null
        || cell.getCellStyle().getIndex() != 0;
  }

  /** Synchronizes table header metadata for every table present in the workbook. */
  public static void syncAllHeaders(XSSFWorkbook workbook) {
    Objects.requireNonNull(workbook, "workbook must not be null");

    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      syncAllHeaders(workbook.getSheetAt(sheetIndex));
    }
  }

  /** Synchronizes table header metadata for sheet tables intersecting the changed cell range. */
  public static void syncAffectedHeaders(XSSFSheet sheet, CellBounds changedRange) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(changedRange, "changedRange must not be null");

    for (XSSFTable table : sheet.getTables()) {
      Optional<CellBounds> headerRange = headerRange(table);
      if (headerRange.isPresent()
          && SheetStructure.intersects(headerRange.orElseThrow(), changedRange)) {
        syncHeader(table);
      }
    }
  }

  private static void syncAllHeaders(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");

    for (XSSFTable table : sheet.getTables()) {
      syncHeader(table);
    }
  }

  private static void syncHeader(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    if (headerRange(table).isPresent()) {
      table.updateHeaders();
    }
  }

  private static Optional<CellBounds> headerRange(XSSFTable table) {
    Objects.requireNonNull(table, "table must not be null");

    if (table.getHeaderRowCount() < 1) {
      return Optional.empty();
    }
    Optional<CellBounds> tableRange =
        SheetStructure.parseOptionalRange(
            Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
    if (tableRange.isEmpty()) {
      return Optional.empty();
    }
    CellBounds resolved = tableRange.orElseThrow();
    return Optional.of(
        new CellBounds(
            resolved.firstRow(),
            Math.min(resolved.lastRow(), resolved.firstRow() + table.getHeaderRowCount() - 1),
            resolved.firstColumn(),
            resolved.lastColumn()));
  }

  /**
   * Rewrites formula cells on one copied sheet from transient cloned table names to final names.
   */
  public static void retargetFormulaCells(
      XSSFSheet sheet, List<String> transientTableNames, List<String> finalTableNames) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(transientTableNames, "transientTableNames must not be null");
    Objects.requireNonNull(finalTableNames, "finalTableNames must not be null");

    Map<String, String> tableNameMapping = tableNameMapping(transientTableNames, finalTableNames);
    if (tableNameMapping.isEmpty()) {
      return;
    }
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA) {
          continue;
        }
        String rewritten = retargetFormula(cell.getCellFormula(), tableNameMapping);
        if (!rewritten.equals(cell.getCellFormula())) {
          FormulaWrites.setRewrittenFormula(
              cell, rewritten, "Table structured-reference retargeting");
        }
      }
    }
  }

  public static String retargetFormula(String formula, Map<String, String> tableNameMapping) {
    Objects.requireNonNull(formula, "formula must not be null");
    Objects.requireNonNull(tableNameMapping, "tableNameMapping must not be null");
    String rewritten = formula;
    for (Map.Entry<String, String> entry : tableNameMapping.entrySet()) {
      rewritten = replaceTableName(rewritten, entry.getKey(), entry.getValue());
    }
    return rewritten;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  private static Map<String, String> tableNameMapping(
      List<String> transientTableNames, List<String> finalTableNames) {
    int mappingCount = Math.min(transientTableNames.size(), finalTableNames.size());
    Map<String, String> mappings = new LinkedHashMap<>();
    for (int index = 0; index < mappingCount; index++) {
      String transientName = requireNonBlank(transientTableNames.get(index), "transientTableNames");
      String finalName = requireNonBlank(finalTableNames.get(index), "finalTableNames");
      if (!transientName.equals(finalName)) {
        mappings.put(transientName, finalName);
      }
    }
    return Map.copyOf(mappings);
  }

  private static String replaceTableName(String formula, String transientName, String finalName) {
    Pattern pattern =
        Pattern.compile("(?<![A-Za-z0-9_\\.])" + Pattern.quote(transientName) + "(?=\\[)");
    Matcher matcher = pattern.matcher(formula);
    return matcher.find() ? matcher.replaceAll(Matcher.quoteReplacement(finalName)) : formula;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not contain null values");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not contain blank values");
    }
    return value;
  }
}
