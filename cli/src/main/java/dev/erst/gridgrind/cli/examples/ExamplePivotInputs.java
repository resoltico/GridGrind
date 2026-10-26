package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import java.util.Optional;

/** Shared pivot-table input builders for shipped example workbook plans. */
final class ExamplePivotInputs {
  private ExamplePivotInputs() {}

  static PivotTableSpec regionalTotalsPivotFromRange(
      String pivotName, String reportSheetName, String sourceSheetName, String sourceRangeAddress) {
    return regionalTotalsPivot(
        pivotName,
        reportSheetName,
        new PivotTableSpec.PivotSourceSpec.Range(
            new SheetName(sourceSheetName), new RangeRef(sourceRangeAddress)));
  }

  static PivotTableSpec regionalTotalsPivotFromTable(
      String pivotName, String reportSheetName, String sourceTableName) {
    return regionalTotalsPivot(
        pivotName,
        reportSheetName,
        new PivotTableSpec.PivotSourceSpec.Table(new TableName(sourceTableName)));
  }

  private static PivotTableSpec regionalTotalsPivot(
      String pivotName, String reportSheetName, PivotTableSpec.PivotSourceSpec source) {
    return new PivotTableSpec(
        pivotName,
        new SheetName(reportSheetName),
        source,
        new PivotTableSpec.Anchor(new CellAddress("A3")),
        List.of("Region"),
        List.of("Stage"),
        List.of(),
        List.of(
            new PivotTableSpec.DataField(
                "Amount",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                Optional.of("#,##0.00"))));
  }
}
