package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MessageInvariant;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.result.ExecutionJournal;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlOpenOptions;
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
            new TableSelector.ByName(new TableName("BudgetTable")),
            "Owner",
            new CellValueSpec.Text(TextSourceSpec.inline("Ada")));
    TableCellSelector.ByColumnName tableCell =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByIndex(new TableSelector.ByName(new TableName("BudgetTable")), 0),
            "Owner");
    ExecutionInputBindings bindings = ExecutionInputBindingsFixtureSupport.bindings(tempDir);
    List<Selector> passThroughSelectors =
        List.of(
            workbookSelector,
            new SheetSelector.All(),
            new SheetSelector.ByName(new SheetName("Budget")),
            new SheetSelector.ByNames(List.of(new SheetName("Budget"), new SheetName("Ops"))),
            new CellSelector.AllUsedInSheet(new SheetName("Budget")),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellSelector.ByAddresses(
                new SheetName("Budget"), List.of(new CellAddress("A1"), new CellAddress("B2"))),
            new RangeSelector.AllOnSheet(new SheetName("Budget")),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new RangeSelector.ByRanges(
                new SheetName("Budget"), List.of(new RangeRef("A1:B2"), new RangeRef("C1:D2"))),
            new RowBandSelector.Span(new SheetName("Budget"), 0, 1),
            new ColumnBandSelector.Insertion(new SheetName("Budget"), 1, 2),
            new DrawingObjectSelector.AllOnSheet(new SheetName("Budget")),
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "Logo"),
            new ChartSelector.AllOnSheet(new SheetName("Budget")),
            new ChartSelector.ByName(new SheetName("Budget"), "OpsChart"),
            new TableSelector.All(),
            new TableSelector.ByName(new TableName("BudgetTable")),
            new TableSelector.ByNames(List.of(new TableName("BudgetTable"))),
            new PivotTableSelector.All(),
            new PivotTableSelector.ByName("BudgetPivot"),
            new PivotTableSelector.ByNames(List.of("BudgetPivot")),
            new NamedRangeSelector.All(),
            new NamedRangeSelector.ByName(new DefinedName("BudgetTotal")),
            new NamedRangeSelector.ByNames(List.of(new DefinedName("BudgetTotal"))),
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.ByName(new DefinedName("BudgetTotal")),
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("OpsTotal"), new SheetName("Ops")))),
            new TableRowSelector.AllRows(new TableSelector.ByName(new TableName("BudgetTable"))),
            new TableRowSelector.ByIndex(new TableSelector.ByName(new TableName("BudgetTable")), 0),
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
            new OoxmlOpenSecuritySpec(Optional.empty())));

    GridGrindRequestDoctor defaultDoctor = new GridGrindRequestDoctor();
    GridGrindRequestDoctor validationOnlyDoctor =
        new GridGrindRequestDoctor(new StaticRequestValidator());

    assertInstanceOf(GridGrindRequestDoctor.class, defaultDoctor);
    assertInstanceOf(GridGrindRequestDoctor.class, validationOnlyDoctor);
    Path tempFile =
        ExecutionContextFixtureSupport.tempFileFactory(
                ExecutionContextFixtureSupport.defaultWorkingDirectory())
            .createTempFile("gridgrind-doctor-", ".xlsx");
    Files.deleteIfExists(tempFile);
  }

  @Test
  void selectorDiagnosticsAndProblemEnrichmentCoverRemainingFallbackBranches() {
    TableCellSelector.ByColumnName tableCell =
        new TableCellSelector.ByColumnName(
            new TableRowSelector.ByIndex(
                new TableSelector.ByNameOnSheet(
                    new TableName("BudgetTable"), new SheetName("Budget")),
                0),
            "Owner");
    TableRowSelector.ByKeyCell keyedRow =
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName(new TableName("BudgetTable")),
            "Owner",
            new CellValueSpec.Text(TextSourceSpec.inline("Ada")));
    List<Selector> summarySelectors =
        List.of(
            new WorkbookSelector.Current(),
            new SheetSelector.All(),
            new SheetSelector.ByName(new SheetName("Budget")),
            new CellSelector.AllUsedInSheet(new SheetName("Budget")),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new RangeSelector.AllOnSheet(new SheetName("Budget")),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("A1"), 2, 2),
            new RowBandSelector.Span(new SheetName("Budget"), 0, 1),
            new RowBandSelector.Insertion(new SheetName("Budget"), 1, 2),
            new ColumnBandSelector.Span(new SheetName("Budget"), 0, 1),
            new ColumnBandSelector.Insertion(new SheetName("Budget"), 1, 2),
            new DrawingObjectSelector.AllOnSheet(new SheetName("Budget")),
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "Logo"),
            new ChartSelector.AllOnSheet(new SheetName("Budget")),
            new ChartSelector.ByName(new SheetName("Budget"), "OpsChart"),
            new TableSelector.All(),
            new TableSelector.ByName(new TableName("BudgetTable")),
            new TableSelector.ByNameOnSheet(new TableName("BudgetTable"), new SheetName("Budget")),
            new PivotTableSelector.All(),
            new PivotTableSelector.ByName("BudgetPivot"),
            new PivotTableSelector.ByNameOnSheet("BudgetPivot", new SheetName("Budget")),
            new NamedRangeSelector.All(),
            new NamedRangeSelector.ByName(new DefinedName("BudgetTotal")),
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
            new NamedRangeSelector.SheetScope(
                new DefinedName("BudgetTotal"), new SheetName("Budget")),
            new TableRowSelector.AllRows(new TableSelector.ByName(new TableName("BudgetTable"))),
            new TableRowSelector.ByIndex(new TableSelector.ByName(new TableName("BudgetTable")), 0),
            keyedRow,
            tableCell);
    List<Selector> expandedSelectors =
        List.of(
            new SheetSelector.ByNames(List.of(new SheetName("Budget"), new SheetName("Ops"))),
            new CellSelector.ByAddresses(
                new SheetName("Budget"), List.of(new CellAddress("A1"), new CellAddress("B2"))),
            new RangeSelector.ByRanges(
                new SheetName("Budget"), List.of(new RangeRef("A1:B2"), new RangeRef("C1:D2"))),
            new TableSelector.ByNames(
                List.of(new TableName("BudgetTable"), new TableName("OpsTable"))),
            new PivotTableSelector.ByNames(List.of("BudgetPivot", "OpsPivot")),
            new NamedRangeSelector.ByNames(
                List.of(new DefinedName("BudgetTotal"), new DefinedName("OpsTotal"))),
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.ByName(new DefinedName("BudgetTotal")),
                    new NamedRangeSelector.WorkbookScope(new DefinedName("GrandTotal")),
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("OpsTotal"), new SheetName("Ops")))));
    GridGrindProblemDetail.Problem problem =
        GridGrindProblems.fromException(
            badRequest(Optional.empty(), Optional.of(11), Optional.empty()),
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"),
                ProblemContextRequestSurfaces.JsonLocation.unavailable()));
    GridGrindProblemDetail.Problem unavailableProblem =
        GridGrindProblems.fromException(
            badRequest(Optional.empty(), Optional.empty(), Optional.of(9)),
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"),
                ProblemContextRequestSurfaces.JsonLocation.unavailable()));
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
        3, ExecutionJournalTargetResolver.expandedTargets(expandedSelectors.get(6)).size());
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
  }

  private static InvalidRequestException badRequest(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    return new InvalidRequestException(
        new MessageInvariant("bad request", jsonPath),
        jsonPath,
        jsonLine,
        jsonColumn,
        new IllegalArgumentException("bad"));
  }
}
