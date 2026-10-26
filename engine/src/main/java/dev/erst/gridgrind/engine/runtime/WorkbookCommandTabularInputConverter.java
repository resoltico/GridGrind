package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;

/** Converts table and pivot-table contract inputs. */
final class WorkbookCommandTabularInputConverter {
  private WorkbookCommandTabularInputConverter() {}

  static ResolvedTable toResolvedTable(TableSpec table) {
    return new ResolvedTable(
        table.name(),
        table.sheetName(),
        table.range(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        table.style(),
        WorkbookCommandSourceSupport.inlineText(table.comment(), "table comment"),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns());
  }

  static ExcelPivotTableDefinition toExcelPivotTableDefinition(PivotTableSpec pivotTable) {
    return new ExcelPivotTableDefinition(
        pivotTable.name(),
        pivotTable.sheetName().value(),
        toExcelPivotTableSource(pivotTable.source()),
        new ExcelPivotTableDefinition.Anchor(pivotTable.anchor().topLeftAddress().value()),
        pivotTable.rowLabels(),
        pivotTable.columnLabels(),
        pivotTable.reportFilters(),
        pivotTable.dataFields().stream()
            .map(WorkbookCommandTabularInputConverter::toExcelPivotTableDataField)
            .toList());
  }

  private static ExcelPivotTableDefinition.Source toExcelPivotTableSource(
      PivotTableSpec.PivotSourceSpec source) {
    return switch (source) {
      case PivotTableSpec.PivotSourceSpec.Range range ->
          new ExcelPivotTableDefinition.Source.Range(
              range.sheetName().value(), range.range().value());
      case PivotTableSpec.PivotSourceSpec.NamedRange namedRange ->
          new ExcelPivotTableDefinition.Source.NamedRange(namedRange.name().value());
      case PivotTableSpec.PivotSourceSpec.Table table ->
          new ExcelPivotTableDefinition.Source.Table(table.name().value());
    };
  }

  private static ExcelPivotTableDefinition.DataField toExcelPivotTableDataField(
      PivotTableSpec.DataField dataField) {
    return new ExcelPivotTableDefinition.DataField(
        dataField.sourceColumnName(),
        dataField.function(),
        dataField.displayName(),
        dataField.valueFormat());
  }
}
