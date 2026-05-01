package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelOoxmlOpenOptions;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Covers remaining executor-side pass-through and fallback branches after DTO tightening. */
class ExecutorCoverageGapsTest {
  @Test
  void passThroughSelectorsAndSecurityDefaultsRemainStable(@TempDir Path tempDir)
      throws IOException {
    WorkbookSelector.Current workbookSelector = new WorkbookSelector.Current();
    TableRowSelector.ByKeyCell keyedRow =
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("BudgetTable"),
            "Owner",
            new CellInput.Text(TextSourceInput.inline("Ada")));
    TableCellSelector.ByColumnName tableCell =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByIndex(new TableSelector.ByName("BudgetTable"), 0), "Owner");
    ExecutionInputBindings bindings = new ExecutionInputBindings(tempDir);
    List<Selector> passThroughSelectors =
        List.of(
            workbookSelector,
            new SheetSelector.All(),
            new SheetSelector.ByName("Budget"),
            new SheetSelector.ByNames(List.of("Budget", "Ops")),
            new CellSelector.AllUsedInSheet("Budget"),
            new CellSelector.ByAddress("Budget", "A1"),
            new CellSelector.ByAddresses("Budget", List.of("A1", "B2")),
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "A1"),
                    new CellSelector.QualifiedAddress("Ops", "B2"))),
            new RangeSelector.AllOnSheet("Budget"),
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new RangeSelector.ByRanges("Budget", List.of("A1:B2", "C1:D2")),
            new RowBandSelector.Span("Budget", 0, 1),
            new ColumnBandSelector.Insertion("Budget", 1, 2),
            new DrawingObjectSelector.AllOnSheet("Budget"),
            new DrawingObjectSelector.ByName("Budget", "Logo"),
            new ChartSelector.AllOnSheet("Budget"),
            new ChartSelector.ByName("Budget", "OpsChart"),
            new TableSelector.All(),
            new TableSelector.ByName("BudgetTable"),
            new TableSelector.ByNames(List.of("BudgetTable")),
            new PivotTableSelector.All(),
            new PivotTableSelector.ByName("BudgetPivot"),
            new PivotTableSelector.ByNames(List.of("BudgetPivot")),
            new NamedRangeSelector.All(),
            new NamedRangeSelector.ByName("BudgetTotal"),
            new NamedRangeSelector.ByNames(List.of("BudgetTotal")),
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.ByName("BudgetTotal"),
                    new NamedRangeSelector.SheetScope("OpsTotal", "Ops"))),
            new TableRowSelector.AllRows(new TableSelector.ByName("BudgetTable")),
            new TableRowSelector.ByIndex(new TableSelector.ByName("BudgetTable"), 0),
            keyedRow,
            tableCell);
    for (Selector selector : passThroughSelectors) {
      assertSame(selector, SourceBackedSelectorResolver.resolve(selector, bindings));
    }

    assertEquals(
        List.of(new ExecutionJournal.Target("WORKBOOK", "Current workbook")),
        ExecutionJournalTargetResolver.expandedTargets(workbookSelector));
    assertInstanceOf(
        ExcelOoxmlOpenOptions.Unencrypted.class,
        OoxmlPackageSecurityConverter.toExcelOpenOptions(
            new OoxmlOpenSecurityInput(Optional.empty())));
    assertTrue(WorkbookCommandStructuredInputConverter.toExcelDifferentialBorder(null).isEmpty());

    GridGrindRequestDoctor defaultDoctor = new GridGrindRequestDoctor();
    GridGrindRequestDoctor validationOnlyDoctor =
        new GridGrindRequestDoctor(new ExecutionValidationSupport());

    assertInstanceOf(GridGrindRequestDoctor.class, defaultDoctor);
    assertInstanceOf(GridGrindRequestDoctor.class, validationOnlyDoctor);
    Path tempFile = GridGrindRequestDoctor.createTempWorkbookFile("gridgrind-doctor-", ".xlsx");
    Files.deleteIfExists(tempFile);
  }

  @Test
  void selectorDiagnosticsAndProblemEnrichmentCoverRemainingFallbackBranches() {
    TableCellSelector.ByColumnName tableCell =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByIndex(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"), 0),
            "Owner");
    TableRowSelector.ByKeyCell keyedRow =
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("BudgetTable"),
            "Owner",
            new CellInput.Text(TextSourceInput.inline("Ada")));
    List<Selector> summarySelectors =
        List.of(
            new WorkbookSelector.Current(),
            new SheetSelector.All(),
            new SheetSelector.ByName("Budget"),
            new CellSelector.AllUsedInSheet("Budget"),
            new CellSelector.ByAddress("Budget", "A1"),
            new RangeSelector.AllOnSheet("Budget"),
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new RangeSelector.RectangularWindow("Budget", "A1", 2, 2),
            new RowBandSelector.Span("Budget", 0, 1),
            new RowBandSelector.Insertion("Budget", 1, 2),
            new ColumnBandSelector.Span("Budget", 0, 1),
            new ColumnBandSelector.Insertion("Budget", 1, 2),
            new DrawingObjectSelector.AllOnSheet("Budget"),
            new DrawingObjectSelector.ByName("Budget", "Logo"),
            new ChartSelector.AllOnSheet("Budget"),
            new ChartSelector.ByName("Budget", "OpsChart"),
            new TableSelector.All(),
            new TableSelector.ByName("BudgetTable"),
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
            new PivotTableSelector.All(),
            new PivotTableSelector.ByName("BudgetPivot"),
            new PivotTableSelector.ByNameOnSheet("BudgetPivot", "Budget"),
            new NamedRangeSelector.All(),
            new NamedRangeSelector.ByName("BudgetTotal"),
            new NamedRangeSelector.WorkbookScope("BudgetTotal"),
            new NamedRangeSelector.SheetScope("BudgetTotal", "Budget"),
            new TableRowSelector.AllRows(new TableSelector.ByName("BudgetTable")),
            new TableRowSelector.ByIndex(new TableSelector.ByName("BudgetTable"), 0),
            keyedRow,
            tableCell);
    List<Selector> expandedSelectors =
        List.of(
            new SheetSelector.ByNames(List.of("Budget", "Ops")),
            new CellSelector.ByAddresses("Budget", List.of("A1", "B2")),
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "A1"),
                    new CellSelector.QualifiedAddress("Ops", "B2"))),
            new RangeSelector.ByRanges("Budget", List.of("A1:B2", "C1:D2")),
            new TableSelector.ByNames(List.of("BudgetTable", "OpsTable")),
            new PivotTableSelector.ByNames(List.of("BudgetPivot", "OpsPivot")),
            new NamedRangeSelector.ByNames(List.of("BudgetTotal", "OpsTotal")),
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.ByName("BudgetTotal"),
                    new NamedRangeSelector.WorkbookScope("GrandTotal"),
                    new NamedRangeSelector.SheetScope("OpsTotal", "Ops"))));
    GridGrindProblemDetail.Problem problem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request", null, 11, null, new IllegalArgumentException("bad")),
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"),
                ProblemContextRequestSurfaces.JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem unavailableProblem =
        GridGrindProblems.fromException(
            new InvalidRequestException(
                "bad request", null, null, 9, new IllegalArgumentException("bad")),
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"),
                ProblemContextRequestSurfaces.JsonLocation.unavailable()));
    DifferentialBorderInput border =
        new DifferentialBorderInput(
            Optional.of(
                new DifferentialBorderSideInput(
                    ExcelBorderStyle.THIN, Optional.of(CellColorReport.rgb("#aabbcc").rgb()))),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    ExcelDifferentialBorder excelBorder =
        WorkbookCommandStructuredInputConverter.toExcelDifferentialBorder(border).orElseThrow();

    for (Selector selector : summarySelectors) {
      assertEquals(
          List.of(ExecutionJournalTargetResolver.summaryTarget(selector)),
          ExecutionJournalTargetResolver.expandedTargets(selector));
    }
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(0)).size());
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(1)).size());
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(2)).size());
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(3)).size());
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(4)).size());
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(5)).size());
    assertEquals(
        2, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(6)).size());
    assertEquals(
        3, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(7)).size());
    assertEquals(Optional.of("Budget"), ExecutionSelectorDiagnosticFields.sheetNameFor(tableCell));
    assertEquals(
        Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(problem).jsonPath());
    assertEquals(
        Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(problem).jsonLine());
    assertEquals(
        Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(problem).jsonColumn());
    assertEquals(
        Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(unavailableProblem)
            .jsonPath());
    assertEquals(
        Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(unavailableProblem)
            .jsonLine());
    assertEquals(
        Optional.empty(),
        DefaultGridGrindRequestExecutorTestSupport.readRequestContext(unavailableProblem)
            .jsonColumn());
    assertEquals(ExcelBorderStyle.THIN, excelBorder.all().style());
  }
}
