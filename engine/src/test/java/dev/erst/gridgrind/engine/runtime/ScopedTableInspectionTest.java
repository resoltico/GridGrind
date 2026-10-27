package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import java.util.List;
import org.junit.jupiter.api.Test;

/** A sheet-scoped selector must not return a same-named object from another sheet. */
class ScopedTableInspectionTest {
  @Test
  void tableInspectionHonorsRequestedSheet() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook
          .getOrCreateSheet("Ops")
          .cells()
          .setRange(
              "A1:B2",
              List.of(
                  List.of(new CellValue.Text("Item"), new CellValue.Text("Value")),
                  List.of(new CellValue.Text("A"), new CellValue.NumberValue(3))));
      workbook
          .tables()
          .setTable(
              new ResolvedTable(
                  new TableName("Queue"),
                  new SheetName("Ops"),
                  new RangeRef("A1:B2"),
                  false,
                  new TableStyleSpec.None()));
      var engine = new WorkbookExecutionEngine();
      var wrongSheet =
          assertInstanceOf(
              WorkbookAssetInspectionResult.TablesResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("wrong"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new TableSelector.ByNameOnSheet(
                                  new TableName("Queue"), new SheetName("Budget")),
                          new WorkbookAssetIntrospectionQuery.GetTables()))
                  .getFirst());
      assertTrue(wrongSheet.tables().isEmpty());
      var matching =
          assertInstanceOf(
              WorkbookAssetInspectionResult.TablesResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("matching"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new TableSelector.ByNameOnSheet(
                                  new TableName("Queue"), new SheetName("Ops")),
                          new WorkbookAssetIntrospectionQuery.GetTables()))
                  .getFirst());
      assertEquals(1, matching.tables().size());
      assertEquals("Ops", matching.tables().getFirst().sheetName().value());
      var insensitive =
          assertInstanceOf(
              WorkbookAssetInspectionResult.TablesResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("insensitive"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new TableSelector.ByNameOnSheet(
                                  new TableName("queue"), new SheetName("ops")),
                          new WorkbookAssetIntrospectionQuery.GetTables()))
                  .getFirst());
      assertEquals(1, insensitive.tables().size());
      var health =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.TableHealthResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("wrong-health"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new TableSelector.ByNameOnSheet(
                                  new TableName("Queue"), new SheetName("Budget")),
                          new dev.erst.gridgrind.core.plan.InspectionAnalysisQuery
                              .AnalyzeTableHealth()))
                  .getFirst());
      assertEquals(0, health.analysis().checkedTableCount());
      var semanticTarget =
          new dev.erst.gridgrind.core.plan.TableCellSelector.ByColumnName(
              new dev.erst.gridgrind.core.plan.TableRowSelector.ByIndex(
                  new TableSelector.ByNameOnSheet(new TableName("queue"), new SheetName("ops")), 0),
              "Value");
      var resolved =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.CellSelector.ByAddress.class,
              new SemanticSelectorResolver(engine)
                  .resolveMutationTarget(
                      workbook,
                      semanticTarget,
                      new dev.erst.gridgrind.core.plan.CellMutationAction.SetCell(
                          new dev.erst.gridgrind.core.model.value.CellValueSpec.NumberValue(7))));
      assertEquals("Ops", resolved.sheetName().value());
      assertEquals("B2", resolved.address().value());
    }
  }

  @Test
  void pivotInspectionHonorsRequestedSheet() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.getOrCreateSheet("Report");
      workbook
          .getOrCreateSheet("Ops")
          .cells()
          .setRange(
              "A1:B2",
              List.of(
                  List.of(new CellValue.Text("Item"), new CellValue.Text("Value")),
                  List.of(new CellValue.Text("A"), new CellValue.NumberValue(3))));
      workbook
          .pivots()
          .setPivotTable(
              new dev.erst.gridgrind.core.model.pivot.PivotTableSpec(
                  "Totals",
                  new SheetName("Report"),
                  new dev.erst.gridgrind.core.model.pivot.PivotTableSpec.PivotSourceSpec.Range(
                      new SheetName("Ops"), new RangeRef("A1:B2")),
                  new dev.erst.gridgrind.core.model.pivot.PivotTableSpec.Anchor(
                      new dev.erst.gridgrind.core.vocabulary.CellAddress("D5")),
                  List.of("Item"),
                  List.of(),
                  List.of(),
                  List.of(
                      new dev.erst.gridgrind.core.model.pivot.PivotTableSpec.DataField(
                          "Value",
                          dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction.SUM,
                          "Total",
                          java.util.Optional.empty()))));
      var engine = new WorkbookExecutionEngine();
      var wrong =
          assertInstanceOf(
              WorkbookAssetInspectionResult.PivotTablesResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("wrong-pivot"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNameOnSheet(
                                  "Totals", new SheetName("Budget")),
                          new WorkbookAssetIntrospectionQuery.GetPivotTables()))
                  .getFirst());
      assertTrue(wrong.pivotTables().isEmpty());
      var matching =
          assertInstanceOf(
              WorkbookAssetInspectionResult.PivotTablesResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("matching-pivot"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNameOnSheet(
                                  "Totals", new SheetName("Report")),
                          new WorkbookAssetIntrospectionQuery.GetPivotTables()))
                  .getFirst());
      assertEquals(1, matching.pivotTables().size());
      var health =
          assertInstanceOf(
              dev.erst.gridgrind.core.plan.WorkbookAnalysisResult.PivotTableHealthResult.class,
              engine
                  .read(
                      workbook,
                      new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
                          new StepId("wrong-pivot-health"),
                          (dev.erst.gridgrind.core.plan.ResolvedTarget)
                              new dev.erst.gridgrind.core.plan.PivotTableSelector.ByNameOnSheet(
                                  "Totals", new SheetName("Budget")),
                          new dev.erst.gridgrind.core.plan.InspectionAnalysisQuery
                              .AnalyzePivotTableHealth()))
                  .getFirst());
      assertEquals(0, health.analysis().checkedPivotTableCount());
    }
  }
}
