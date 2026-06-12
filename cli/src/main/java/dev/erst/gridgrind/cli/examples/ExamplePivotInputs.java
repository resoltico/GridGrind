package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import java.util.Optional;

/** Shared pivot-table input builders for shipped example workbook plans. */
final class ExamplePivotInputs {
  private ExamplePivotInputs() {}

  static PivotTableInput regionalTotalsPivotFromRange(
      String pivotName, String reportSheetName, String sourceSheetName, String sourceRangeAddress) {
    return regionalTotalsPivot(
        pivotName,
        reportSheetName,
        new PivotTableInput.Source.Range(sourceSheetName, sourceRangeAddress));
  }

  static PivotTableInput regionalTotalsPivotFromTable(
      String pivotName, String reportSheetName, String sourceTableName) {
    return regionalTotalsPivot(
        pivotName, reportSheetName, new PivotTableInput.Source.Table(sourceTableName));
  }

  private static PivotTableInput regionalTotalsPivot(
      String pivotName, String reportSheetName, PivotTableInput.Source source) {
    return new PivotTableInput(
        pivotName,
        reportSheetName,
        source,
        new PivotTableInput.Anchor("A3"),
        List.of("Region"),
        List.of("Stage"),
        List.of(),
        List.of(
            new PivotTableInput.DataField(
                "Amount",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                Optional.of("#,##0.00"))));
  }
}
