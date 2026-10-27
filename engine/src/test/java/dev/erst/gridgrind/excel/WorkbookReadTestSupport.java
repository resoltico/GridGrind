package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.model.value.CellReport;
import dev.erst.gridgrind.core.plan.CellReadProjection;
import dev.erst.gridgrind.core.plan.ResolvedInspectionStep;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;

/** Shared defaults for read-command and read-result tests after projection became explicit. */
final class WorkbookReadTestSupport {
  private WorkbookReadTestSupport() {}

  static CellReadProjection projection() {
    return CellReadProjection.defaults();
  }

  static ResolvedInspectionStep getCells(String stepId, String sheetName, List<String> addresses) {
    return new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
        new StepId(stepId),
        new dev.erst.gridgrind.core.plan.CellSelector.ByAddresses(
            new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
            addresses.stream().map(dev.erst.gridgrind.core.vocabulary.CellAddress::new).toList()),
        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetCells(
            java.util.Optional.of(projection())));
  }

  static ResolvedInspectionStep getWindow(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
        new StepId(stepId),
        new dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow(
            new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
            new dev.erst.gridgrind.core.vocabulary.CellAddress(topLeftAddress),
            rowCount,
            columnCount),
        new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetWindow(
            java.util.Optional.of(projection()), false));
  }

  static ResolvedInspectionStep getSheetSchema(
      String stepId, String sheetName, String topLeftAddress, int rowCount, int columnCount) {
    return new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
        new StepId(stepId),
        new dev.erst.gridgrind.core.plan.RangeSelector.RectangularWindow(
            new dev.erst.gridgrind.core.vocabulary.SheetName(sheetName),
            new dev.erst.gridgrind.core.vocabulary.CellAddress(topLeftAddress),
            rowCount,
            columnCount),
        new dev.erst.gridgrind.core.plan.InspectionSurfaceQuery.GetSheetSchema(
            java.util.Optional.of(projection())));
  }

  static dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult cellsResult(
      String stepId, String sheetName, List<CellReport> cells) {
    return new dev.erst.gridgrind.core.plan.SheetInspectionResult.CellsResult(
        new StepId(stepId), new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), cells);
  }

  static dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.SheetSchemaResult
      sheetSchemaResult(String stepId, SheetSchemaReport surface) {
    return new dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult.SheetSchemaResult(
        new StepId(stepId), surface);
  }
}
