package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableColumnDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;

/** Converts table and pivot-table contract inputs. */
final class WorkbookCommandTabularInputConverter {
  private WorkbookCommandTabularInputConverter() {}

  static ExcelTableDefinition toExcelTableDefinition(TableInput table) {
    return new ExcelTableDefinition(
        table.name(),
        table.sheetName(),
        table.range(),
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

  static ExcelPivotTableDefinition toExcelPivotTableDefinition(PivotTableInput pivotTable) {
    return new ExcelPivotTableDefinition(
        pivotTable.name(),
        pivotTable.sheetName(),
        toExcelPivotTableSource(pivotTable.source()),
        new ExcelPivotTableDefinition.Anchor(pivotTable.anchor().topLeftAddress()),
        pivotTable.rowLabels(),
        pivotTable.columnLabels(),
        pivotTable.reportFilters(),
        pivotTable.dataFields().stream()
            .map(WorkbookCommandTabularInputConverter::toExcelPivotTableDataField)
            .toList());
  }

  static ExcelTableStyle toExcelTableStyle(TableStyleInput style) {
    return switch (style) {
      case TableStyleInput.None _ -> new ExcelTableStyle.None();
      case TableStyleInput.Named named ->
          new ExcelTableStyle.Named(
              named.name(),
              named.showFirstColumn(),
              named.showLastColumn(),
              named.showRowStripes(),
              named.showColumnStripes());
    };
  }

  private static ExcelPivotTableDefinition.Source toExcelPivotTableSource(
      PivotTableInput.Source source) {
    return switch (source) {
      case PivotTableInput.Source.Range range ->
          new ExcelPivotTableDefinition.Source.Range(range.sheetName(), range.range());
      case PivotTableInput.Source.NamedRange namedRange ->
          new ExcelPivotTableDefinition.Source.NamedRange(namedRange.name());
      case PivotTableInput.Source.Table table ->
          new ExcelPivotTableDefinition.Source.Table(table.name());
    };
  }

  private static ExcelPivotTableDefinition.DataField toExcelPivotTableDataField(
      PivotTableInput.DataField dataField) {
    return new ExcelPivotTableDefinition.DataField(
        dataField.sourceColumnName(),
        dataField.function(),
        dataField.displayName(),
        dataField.valueFormat());
  }
}
