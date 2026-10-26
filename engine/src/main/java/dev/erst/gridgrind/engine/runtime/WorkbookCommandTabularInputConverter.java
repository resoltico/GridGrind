package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.excel.ExcelTableColumnDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;

/** Converts table and pivot-table contract inputs. */
final class WorkbookCommandTabularInputConverter {
  private WorkbookCommandTabularInputConverter() {}

  static ExcelTableDefinition toExcelTableDefinition(TableSpec table) {
    return new ExcelTableDefinition(
        table.name().value(),
        table.sheetName().value(),
        table.range().value(),
        table.showTotalsRow(),
        table.hasAutofilter(),
        toExcelTableStyle(table.style()),
        WorkbookCommandSourceSupport.inlineText(table.comment(), "table comment"),
        table.published(),
        table.insertRow(),
        table.insertRowShift(),
        table.headerRowCellStyle(),
        table.dataCellStyle(),
        table.totalsRowCellStyle(),
        table.columns().stream()
            .map(
                column ->
                    new ExcelTableColumnDefinition(
                        column.columnIndex(),
                        column.uniqueName(),
                        column.totalsRowLabel(),
                        column.totalsRowFunction(),
                        column.calculatedColumnFormula()))
            .toList());
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

  static ExcelTableStyle toExcelTableStyle(TableStyleSpec style) {
    return switch (style) {
      case TableStyleSpec.None _ -> new ExcelTableStyle.None();
      case TableStyleSpec.Named named ->
          new ExcelTableStyle.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
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
