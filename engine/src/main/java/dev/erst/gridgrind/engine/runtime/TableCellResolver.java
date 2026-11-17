package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellReadFacet;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfSheet;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Resolves exact cells within observed native tables using canonical geometry and row keys. */
final class TableCellResolver {
  private static final String RESOLUTION_STEP_ID = "__semantic-selector-resolution__";
  private final WorkbookExecutionEngine workbookEngine;

  TableCellResolver(WorkbookExecutionEngine workbookEngine) {
    this.workbookEngine = Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
  }

  ResolvedTableCell resolve(
      XssfWorkbook workbook, TableCellSelector.ByColumnName selector, boolean allowZeroMatch) {
    ResolvedTableRow resolvedRow = resolveTableRow(workbook, selector.row(), allowZeroMatch);
    int columnOffset = tableColumnOffset(resolvedRow.table(), selector.columnName());
    if (!resolvedRow.matched()) {
      TableRowSelector.ByKeyCell byKeyCell = (TableRowSelector.ByKeyCell) selector.row();
      return new ResolvedTableCell(
          resolvedRow.table().sheetName().value(),
          resolvedRow.table().name().value(),
          byKeyCell.columnName(),
          null);
    }
    return new ResolvedTableCell(
        resolvedRow.table().sheetName().value(),
        resolvedRow.table().name().value(),
        selector.columnName(),
        new CellAddress(
            Objects.requireNonNull(resolvedRow.rowIndex(), "resolved row index must be present"),
            resolvedRow.firstColumnIndex() + columnOffset));
  }

  private ResolvedTableRow resolveTableRow(
      XssfWorkbook workbook, TableRowSelector selector, boolean allowZeroMatch) {
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

  @SuppressWarnings(
      "PMD.AvoidInstantiatingObjectsInLoops") // Each row needs its immutable coordinate; only one
  // snapshot is retained at a time.
  private ResolvedTableRow resolveKeyedTableRow(
      XssfWorkbook workbook,
      ResolvedTable table,
      TableRowSelector.ByKeyCell selector,
      boolean allowZeroMatch) {
    int keyColumnOffset = tableColumnOffset(table.table(), selector.columnName());
    XssfSheet sheet = workbook.sheet(table.table().sheetName().value());
    Integer matchedRowIndex = null;
    var projection = CellReadProjection.of(CellReadFacet.VALUE, CellReadFacet.FORMULA);
    for (int rowIndex = table.dataFirstRowIndex();
        rowIndex <= table.dataLastRowIndex();
        rowIndex++) {
      var address = new CellAddress(rowIndex, table.firstColumnIndex() + keyColumnOffset);
      CellReport snapshot = sheet.cells().snapshotCell(address, projection);
      if (matchesKeyCell(snapshot, selector.expectedValue())) {
        if (matchedRowIndex != null) {
          throw new IllegalArgumentException(
              "table row selector matched more than one row for table "
                  + table.table().name().value()
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
              + table.table().name().value()
              + " column "
              + selector.columnName());
    }
    return new ResolvedTableRow(table.table(), matchedRowIndex, table.firstColumnIndex());
  }

  static boolean matchesKeyCell(CellReport snapshot, CellValueSpec expectedValue) {
    var report = snapshot;
    return switch (expectedValue) {
      case CellValueSpec.Blank _ ->
          report instanceof dev.erst.gridgrind.core.model.value.CellReport.BlankReport;
      case CellValueSpec.Text text ->
          report instanceof dev.erst.gridgrind.core.model.value.CellReport.TextReport textReport
              && textReport
                  .value()
                  .map(dev.erst.gridgrind.core.model.value.CellValue.Text::value)
                  .orElseThrow()
                  .equals(inlineText(text.value(), "table row key TEXT"));
      case CellValueSpec.NumberValue numberValue ->
          report instanceof dev.erst.gridgrind.core.model.value.CellReport.NumberReport numberReport
              && Double.compare(
                      numberReport
                          .value()
                          .map(dev.erst.gridgrind.core.model.value.CellValue.NumberValue::value)
                          .orElseThrow(),
                      numberValue.value())
                  == 0;
      case CellValueSpec.BooleanValue booleanValue ->
          report
                  instanceof
                  dev.erst.gridgrind.core.model.value.CellReport.BooleanReport booleanReport
              && booleanReport
                  .value()
                  .map(dev.erst.gridgrind.core.model.value.CellValue.BooleanValue::value)
                  .orElseThrow()
                  .equals(booleanValue.value());
      case CellValueSpec.Formula formula ->
          report
                  instanceof
                  dev.erst.gridgrind.core.model.value.CellReport.FormulaReport formulaReport
              && formulaReport
                  .formula()
                  .orElseThrow()
                  .equals(inlineText(formula.value(), "table row key FORMULA"));
      default -> false;
    };
  }

  private static String inlineText(
      dev.erst.gridgrind.core.source.TextSourceSpec source, String context) {
    if (source instanceof dev.erst.gridgrind.core.source.TextSourceSpec.Inline inline) {
      return inline.text();
    }
    throw new IllegalStateException(context + " must be resolved to INLINE text before execution");
  }

  private ResolvedTable resolveExactTable(XssfWorkbook workbook, TableSelector selector) {
    dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult result =
        (dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult.TablesResult)
            workbookEngine
                .read(
                    workbook,
                    new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                        new StepId(RESOLUTION_STEP_ID),
                        (dev.erst.gridgrind.core.plan.ResolvedTarget) selector,
                        new dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery
                            .GetTables()))
                .getFirst();
    List<TableEntryReport> tables = result.tables();
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
    TableEntryReport table = tables.getFirst();
    RangeRef bounds = RangeRef.parse(table.range().value());
    return new ResolvedTable(
        table,
        bounds.firstColumn(),
        bounds.firstRow() + table.structure().headerRowCount(),
        bounds.lastRow() - table.structure().totalsRowCount());
  }

  private static int tableColumnOffset(TableEntryReport table, String columnName) {
    String lookup = columnName.toUpperCase(Locale.ROOT);
    for (int index = 0; index < table.structure().columnNames().size(); index++) {
      if (table.structure().columnNames().get(index).toUpperCase(Locale.ROOT).equals(lookup)) {
        return index;
      }
    }
    throw new IllegalArgumentException(
        "table " + table.name().value() + " does not contain column " + columnName);
  }

  private static TableSelector tableSelectorFor(TableRowSelector selector) {
    return selector instanceof TableRowSelector.ByIndex byIndex
        ? byIndex.table()
        : ((TableRowSelector.ByKeyCell) selector).table();
  }

  private record ResolvedTable(
      TableEntryReport snapshot,
      int firstColumnIndex,
      int dataFirstRowIndex,
      int dataLastRowIndex) {
    TableEntryReport table() {
      return snapshot;
    }

    int dataRowCount() {
      return Math.max(0, dataLastRowIndex - dataFirstRowIndex + 1);
    }
  }

  private record ResolvedTableRow(
      TableEntryReport table, @Nullable Integer rowIndex, int firstColumnIndex) {
    boolean matched() {
      return rowIndex != null;
    }
  }

  record ResolvedTableCell(
      String sheetName, String tableName, String lookupColumnName, @Nullable CellAddress address) {
    boolean matched() {
      return address != null;
    }
  }
}
