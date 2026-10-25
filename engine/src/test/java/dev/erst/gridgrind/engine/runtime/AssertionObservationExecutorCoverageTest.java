package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Covers edge-case observation results independently from assertion orchestration. */
class AssertionObservationExecutorCoverageTest {
  @Test
  void observationHelperBranchesRejectUnsupportedTargetsAndReturnZeroMatches() throws Exception {
    WorkbookExecutionEngine readExecutor = new WorkbookExecutionEngine();
    AssertionObservationExecutor observationExecutor =
        new AssertionObservationExecutor(readExecutor, new SemanticSelectorResolver(readExecutor));

    IllegalArgumentException unsupportedPresence =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                observationExecutor.presenceObservation(
                    new StepId("presence"), new WorkbookSelector.Current(), null, null));
    assertTrue(unsupportedPresence.getMessage().contains("Unsupported presence assertion target"));

    NullPointerException nullWorkbook =
        assertThrows(
            NullPointerException.class,
            () ->
                observationExecutor.executeObservation(
                    new StepId("charts"),
                    new WorkbookSelector.Current(),
                    new WorkbookAssetIntrospectionQuery.GetCharts(),
                    null,
                    null));
    assertEquals("workbook must not be null", nullWorkbook.getMessage());

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      IllegalArgumentException unsupportedCharts =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  observationExecutor.executeObservation(
                      new StepId("charts"),
                      new WorkbookSelector.Current(),
                      new WorkbookAssetIntrospectionQuery.GetCharts(),
                      workbook,
                      new WorkbookLocation.UnsavedWorkbook()));
      assertEquals("Unsupported chart inspection target", unsupportedCharts.getMessage());
    }

    IllegalArgumentException unsupportedObservedCount =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                AssertionObservationExecutor.observedCount(
                    new WorkbookInspectionResult.WorkbookSummaryResult(
                        new StepId("summary"),
                        new WorkbookSummary.WithSheets(
                            1,
                            List.of(new SheetLabel("Budget")),
                            new SheetLabel("Budget"),
                            List.of(new SheetLabel("Budget")),
                            0,
                            false))));
    assertTrue(
        unsupportedObservedCount.getMessage().contains("Unsupported presence observation result"));
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookInspectionResult.NamedRangesResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    new StepId("missing-range"),
                    new NamedRangeSelector.WorkbookScope("MissingTotal")))
            .namedRanges());
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookAssetInspectionResult.TablesResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    new StepId("missing-table"), new TableSelector.ByName("MissingTable")))
            .tables());
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookAssetInspectionResult.PivotTablesResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    new StepId("missing-pivot"), new PivotTableSelector.ByName("Missing Pivot")))
            .pivotTables());
    assertEquals(
        "MissingSheet",
        assertInstanceOf(
                WorkbookAssetInspectionResult.ChartsResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    new StepId("missing-chart-sheet"),
                    new ChartSelector.AllOnSheet(new SheetName("MissingSheet"))))
            .sheetName());
    assertEquals(
        "MissingSheet",
        assertInstanceOf(
                WorkbookAssetInspectionResult.ChartsResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    new StepId("missing-chart-name"),
                    new ChartSelector.ByName(new SheetName("MissingSheet"), "MissingChart")))
            .sheetName());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            AssertionObservationExecutor.zeroMatchPresenceObservation(
                new StepId("unsupported-zero-match"), new WorkbookSelector.Current()));
  }
}
