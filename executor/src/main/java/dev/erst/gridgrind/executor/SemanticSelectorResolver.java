package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelSheet;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.WorkbookReadResult;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Resolves semantic selectors into exact executable targets without inventing new semantics. */
final class SemanticSelectorResolver {
  private static final String RESOLUTION_STEP_ID = "__semantic-selector-resolution__";
  private static final Set<Class<? extends MutationAction>> EXACT_CELL_MUTATION_ACTIONS =
      Set.of(
          MutationAction.SetCell.class,
          MutationAction.SetHyperlink.class,
          MutationAction.ClearHyperlink.class,
          MutationAction.SetComment.class,
          MutationAction.ClearComment.class);
  private static final Set<Class<? extends InspectionQuery>> EXACT_CELL_INSPECTION_QUERIES =
      Set.of(
          InspectionQuery.GetCells.class,
          InspectionQuery.GetHyperlinks.class,
          InspectionQuery.GetComments.class);
  private static final Set<Class<? extends Assertion>> EXACT_CELL_ASSERTIONS =
      Set.of(
          Assertion.CellValue.class,
          Assertion.DisplayValue.class,
          Assertion.FormulaText.class,
          Assertion.CellStyle.class);

  private final WorkbookReadExecutor readExecutor;

  SemanticSelectorResolver(WorkbookReadExecutor readExecutor) {
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
  }

  Selector resolveMutationTarget(ExcelWorkbook workbook, Selector target, MutationAction action) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(action, "action must not be null");
    return EXACT_CELL_MUTATION_ACTIONS.contains(action.getClass())
        ? resolveExactCellTarget(workbook, target)
        : target;
  }

  ResolvedInspectionTarget resolveInspectionTarget(
      String stepId, ExcelWorkbook workbook, Selector target, InspectionQuery query) {
    Objects.requireNonNull(stepId, "stepId must not be null");
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(query, "query must not be null");
    return EXACT_CELL_INSPECTION_QUERIES.contains(query.getClass())
        ? resolveExactCellInspectionTarget(stepId, workbook, target, query)
        : ResolvedInspectionTarget.direct(target);
  }

  Selector resolveAssertionTarget(ExcelWorkbook workbook, Selector target, Assertion assertion) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(assertion, "assertion must not be null");
    return EXACT_CELL_ASSERTIONS.contains(assertion.getClass())
        ? resolveExactCellTarget(workbook, target)
        : target;
  }

  private Selector resolveExactCellTarget(ExcelWorkbook workbook, Selector target) {
    if (target instanceof TableCellSelector.ByColumnName tableCell) {
      ResolvedTableCell resolvedCell = resolveTableCell(workbook, tableCell, false);
      return new CellSelector.ByAddress(resolvedCell.sheetName(), resolvedCell.address());
    }
    return target;
  }

  private ResolvedInspectionTarget resolveExactCellInspectionTarget(
      String stepId, ExcelWorkbook workbook, Selector target, InspectionQuery query) {
    if (!(target instanceof TableCellSelector.ByColumnName tableCell)) {
      return ResolvedInspectionTarget.direct(target);
    }
    ResolvedTableCell resolvedCell = resolveTableCell(workbook, tableCell, true);
    if (!resolvedCell.matched()) {
      return ResolvedInspectionTarget.shortCircuit(
          emptyInspectionResult(stepId, resolvedCell.sheetName(), query));
    }
    return ResolvedInspectionTarget.direct(
        new CellSelector.ByAddress(resolvedCell.sheetName(), resolvedCell.address()));
  }

  private ResolvedTableCell resolveTableCell(
      ExcelWorkbook workbook, TableCellSelector.ByColumnName selector, boolean allowZeroMatch) {
    ResolvedTableRow resolvedRow = resolveTableRow(workbook, selector.row(), allowZeroMatch);
    int columnOffset = tableColumnOffset(resolvedRow.table(), selector.columnName());
    if (!resolvedRow.matched()) {
      TableRowSelector.ByKeyCell byKeyCell = (TableRowSelector.ByKeyCell) selector.row();
      return new ResolvedTableCell(
          resolvedRow.table().sheetName(),
          resolvedRow.table().name(),
          byKeyCell.columnName(),
          null);
    }
    return new ResolvedTableCell(
        resolvedRow.table().sheetName(),
        resolvedRow.table().name(),
        selector.columnName(),
        a1Address(resolvedRow.rowIndex(), resolvedRow.firstColumnIndex() + columnOffset));
  }

  private ResolvedTableRow resolveTableRow(
      ExcelWorkbook workbook, TableRowSelector selector, boolean allowZeroMatch) {
    ResolvedTable table = resolveExactTable(workbook, tableSelectorFor(selector));
    int dataRowCount = table.dataRowCount();
    if (selector instanceof TableRowSelector.ByIndex byIndex) {
      if (byIndex.rowIndex() >= dataRowCount) {
        throw new IllegalArgumentException(
            "table row index "
                + byIndex.rowIndex()
                + " is outside the data-row bounds for table "
                + table.snapshot().name()
                + " (rowCount="
                + dataRowCount
                + ")");
      }
      return new ResolvedTableRow(
          table.snapshot(),
          table.dataFirstRowIndex() + byIndex.rowIndex(),
          table.firstColumnIndex());
    }
    return resolveKeyedTableRow(
        workbook, table, (TableRowSelector.ByKeyCell) selector, allowZeroMatch);
  }

  private ResolvedTableRow resolveKeyedTableRow(
      ExcelWorkbook workbook,
      ResolvedTable table,
      TableRowSelector.ByKeyCell selector,
      boolean allowZeroMatch) {
    int keyColumnOffset = tableColumnOffset(table.table(), selector.columnName());
    ExcelSheet sheet = workbook.sheet(table.table().sheetName());
    Integer matchedRowIndex = null;
    for (int rowIndex = table.dataFirstRowIndex();
        rowIndex <= table.dataLastRowIndex();
        rowIndex++) {
      String address = a1Address(rowIndex, table.firstColumnIndex() + keyColumnOffset);
      ExcelCellSnapshot snapshot = sheet.snapshotCell(address);
      if (matchesKeyCell(snapshot, selector.expectedValue())) {
        if (matchedRowIndex != null) {
          throw new IllegalArgumentException(
              "table row selector matched more than one row for table "
                  + table.table().name()
                  + " column "
                  + selector.columnName());
        }
        matchedRowIndex = rowIndex;
      }
    }
    if (matchedRowIndex == null) {
      if (allowZeroMatch) {
        return new ResolvedTableRow(table.table(), null, table.firstColumnIndex());
      }
      throw new IllegalArgumentException(
          "table row selector matched no rows for table "
              + table.table().name()
              + " column "
              + selector.columnName());
    }
    return new ResolvedTableRow(table.table(), matchedRowIndex, table.firstColumnIndex());
  }

  private boolean matchesKeyCell(ExcelCellSnapshot snapshot, CellInput expectedValue) {
    dev.erst.gridgrind.contract.dto.CellReport report =
        InspectionResultCellReportSupport.toCellReport(snapshot);
    if (expectedValue instanceof CellInput.Blank) {
      return report instanceof dev.erst.gridgrind.contract.dto.CellReport.BlankReport;
    }
    if (expectedValue instanceof CellInput.Text text) {
      return report instanceof dev.erst.gridgrind.contract.dto.CellReport.TextReport textReport
          && textReport.stringValue().equals(inlineText(text.source(), "table row key TEXT"));
    }
    if (expectedValue instanceof CellInput.Numeric numeric) {
      return report instanceof dev.erst.gridgrind.contract.dto.CellReport.NumberReport numberReport
          && Double.compare(numberReport.numberValue(), numeric.number()) == 0;
    }
    if (expectedValue instanceof CellInput.BooleanValue booleanValue) {
      return report
              instanceof dev.erst.gridgrind.contract.dto.CellReport.BooleanReport booleanReport
          && booleanReport.booleanValue().equals(booleanValue.bool());
    }
    CellInput.Formula formula = (CellInput.Formula) expectedValue;
    return report instanceof dev.erst.gridgrind.contract.dto.CellReport.FormulaReport formulaReport
        && formulaReport.formula().equals(inlineText(formula.source(), "table row key FORMULA"));
  }

  private static String inlineText(
      dev.erst.gridgrind.contract.source.TextSourceInput source, String context) {
    if (source instanceof dev.erst.gridgrind.contract.source.TextSourceInput.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(context + " must be resolved to INLINE text before execution");
  }

  private ResolvedTable resolveExactTable(ExcelWorkbook workbook, TableSelector selector) {
    WorkbookReadResult.TablesResult result =
        (WorkbookReadResult.TablesResult)
            readExecutor
                .apply(
                    workbook,
                    new WorkbookReadCommand.GetTables(
                        RESOLUTION_STEP_ID, SelectorConverter.toExcelTableSelection(selector)))
                .getFirst();
    List<ExcelTableSnapshot> tables = result.tables();
    if (tables.isEmpty()) {
      if (selector instanceof TableSelector.ByName byName) {
        throw new IllegalArgumentException("table not found: " + byName.name());
      }
      TableSelector.ByNameOnSheet byNameOnSheet = (TableSelector.ByNameOnSheet) selector;
      throw new IllegalArgumentException(
          "table not found on expected sheet: "
              + byNameOnSheet.name()
              + "@"
              + byNameOnSheet.sheetName());
    }
    ExcelTableSnapshot table = tables.getFirst();
    if (selector instanceof TableSelector.ByNameOnSheet byNameOnSheet
        && !table.sheetName().equals(byNameOnSheet.sheetName())) {
      throw new IllegalArgumentException(
          "table not found on expected sheet: "
              + byNameOnSheet.name()
              + "@"
              + byNameOnSheet.sheetName());
    }
    RangeBounds bounds = RangeBounds.parse(table.range());
    return new ResolvedTable(
        table,
        bounds.firstColumn(),
        bounds.firstRow() + table.headerRowCount(),
        bounds.lastRow() - table.totalsRowCount());
  }

  private static int tableColumnOffset(ExcelTableSnapshot table, String columnName) {
    String lookup = columnName.toUpperCase(Locale.ROOT);
    for (int index = 0; index < table.columnNames().size(); index++) {
      if (table.columnNames().get(index).toUpperCase(Locale.ROOT).equals(lookup)) {
        return index;
      }
    }
    throw new IllegalArgumentException(
        "table " + table.name() + " does not contain column " + columnName);
  }

  private static TableSelector tableSelectorFor(TableRowSelector selector) {
    return selector instanceof TableRowSelector.ByIndex byIndex
        ? byIndex.table()
        : ((TableRowSelector.ByKeyCell) selector).table();
  }

  private static String a1Address(int rowIndex, int columnIndex) {
    return columnLabel(columnIndex) + (rowIndex + 1);
  }

  private static String columnLabel(int columnIndex) {
    int value = columnIndex + 1;
    StringBuilder builder = new StringBuilder();
    while (value > 0) {
      int remainder = (value - 1) % 26;
      builder.append((char) ('A' + remainder));
      value = (value - 1) / 26;
    }
    return builder.reverse().toString();
  }

  private static InspectionResult emptyInspectionResult(
      String stepId, String sheetName, InspectionQuery query) {
    if (query instanceof InspectionQuery.GetCells) {
      return new InspectionResult.CellsResult(stepId, sheetName, List.of());
    }
    if (query instanceof InspectionQuery.GetHyperlinks) {
      return new InspectionResult.HyperlinksResult(stepId, sheetName, List.of());
    }
    return new InspectionResult.CommentsResult(stepId, sheetName, List.of());
  }

  record ResolvedInspectionTarget(Selector selector, InspectionResult shortCircuitResult) {
    ResolvedInspectionTarget {
      if (selector == null && shortCircuitResult == null) {
        throw new IllegalArgumentException(
            "resolved inspection target requires either a selector or a short-circuit result");
      }
    }

    static ResolvedInspectionTarget direct(Selector selector) {
      return new ResolvedInspectionTarget(
          Objects.requireNonNull(selector, "selector must not be null"), null);
    }

    static ResolvedInspectionTarget shortCircuit(InspectionResult result) {
      return new ResolvedInspectionTarget(
          null, Objects.requireNonNull(result, "result must not be null"));
    }

    boolean isShortCircuit() {
      return shortCircuitResult != null;
    }
  }

  private record ResolvedTable(
      ExcelTableSnapshot snapshot,
      int firstColumnIndex,
      int dataFirstRowIndex,
      int dataLastRowIndex) {
    ExcelTableSnapshot table() {
      return snapshot;
    }

    int dataRowCount() {
      return Math.max(0, dataLastRowIndex - dataFirstRowIndex + 1);
    }
  }

  private record ResolvedTableRow(
      ExcelTableSnapshot table, Integer rowIndex, int firstColumnIndex) {
    boolean matched() {
      return rowIndex != null;
    }
  }

  private record ResolvedTableCell(
      String sheetName, String tableName, String lookupColumnName, String address) {
    boolean matched() {
      return address != null;
    }
  }

  private record RangeBounds(int firstRow, int lastRow, int firstColumn, int lastColumn) {
    static RangeBounds parse(String range) {
      String[] parts = Objects.requireNonNull(range, "range must not be null").split(":", -1);
      CellAddress first = CellAddress.parse(parts[0]);
      CellAddress last = CellAddress.parse(parts[parts.length - 1]);
      return new RangeBounds(
          Math.min(first.rowIndex(), last.rowIndex()),
          Math.max(first.rowIndex(), last.rowIndex()),
          Math.min(first.columnIndex(), last.columnIndex()),
          Math.max(first.columnIndex(), last.columnIndex()));
    }
  }

  private record CellAddress(int rowIndex, int columnIndex) {
    static CellAddress parse(String address) {
      String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
      String columnText = normalized.replaceAll("[0-9].*$", "");
      int columnIndex = 0;
      for (int index = 0; index < columnText.length(); index++) {
        columnIndex = (columnIndex * 26) + (columnText.charAt(index) - 'A' + 1);
      }
      int rowIndex = Integer.parseInt(normalized.substring(columnText.length())) - 1;
      return new CellAddress(rowIndex, columnIndex - 1);
    }
  }
}
