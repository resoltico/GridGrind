package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.WorkbookSummary;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
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
                    "presence", new WorkbookSelector.Current(), null, null));
    assertTrue(unsupportedPresence.getMessage().contains("Unsupported presence assertion target"));

    NullPointerException nullWorkbook =
        assertThrows(
            NullPointerException.class,
            () ->
                observationExecutor.executeObservation(
                    "charts",
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
                      "charts",
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
                        "summary",
                        new WorkbookSummary.WithSheets(
                            1, List.of("Budget"), "Budget", List.of("Budget"), 0, false))));
    assertTrue(
        unsupportedObservedCount.getMessage().contains("Unsupported presence observation result"));
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookInspectionResult.NamedRangesResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    "missing-range", new NamedRangeSelector.WorkbookScope("MissingTotal")))
            .namedRanges());
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookAssetInspectionResult.TablesResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    "missing-table", new TableSelector.ByName("MissingTable")))
            .tables());
    assertEquals(
        List.of(),
        assertInstanceOf(
                WorkbookAssetInspectionResult.PivotTablesResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    "missing-pivot", new PivotTableSelector.ByName("Missing Pivot")))
            .pivotTables());
    assertEquals(
        "MissingSheet",
        assertInstanceOf(
                WorkbookAssetInspectionResult.ChartsResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    "missing-chart-sheet", new ChartSelector.AllOnSheet("MissingSheet")))
            .sheetName());
    assertEquals(
        "MissingSheet",
        assertInstanceOf(
                WorkbookAssetInspectionResult.ChartsResult.class,
                AssertionObservationExecutor.zeroMatchPresenceObservation(
                    "missing-chart-name", new ChartSelector.ByName("MissingSheet", "MissingChart")))
            .sheetName());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            AssertionObservationExecutor.zeroMatchPresenceObservation(
                "unsupported-zero-match", new WorkbookSelector.Current()));
  }
}
