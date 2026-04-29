package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCategory;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
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
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;

/** Coverage tests for the execution-journal helpers added in Phase 5. */
class ExecutionJournalCoverageTest {
  @Test
  void defaultExecutorMethodForwardsAndRejectsNullSink() {
    WorkbookPlan request =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of());
    GridGrindResponse.Success expected =
        GridGrindResponses.success(
            null,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of());
    AtomicBoolean called = new AtomicBoolean(false);
    GridGrindRequestExecutor executor =
        (ignoredRequest, ignoredBindings, ignoredSink) -> {
          called.set(true);
          return expected;
        };

    assertSame(expected, executor.execute(request, ExecutionJournalSink.NOOP));
    assertTrue(called.get());
    assertThrows(
        NullPointerException.class, () -> executor.execute(request, (ExecutionJournalSink) null));
  }

  @Test
  void targetResolverSummarizesAndExpandsSelectorFamilies() {
    InspectionStep workbookSummary =
        new InspectionStep(
            "step-1", new WorkbookSelector.Current(), new InspectionQuery.GetWorkbookSummary());
    assertEquals(
        List.of(new ExecutionJournal.Target("WORKBOOK", "Current workbook")),
        ExecutionJournalTargetResolver.resolve(workbookSummary, ExecutionJournalLevel.SUMMARY));
    assertEquals(
        List.of(new ExecutionJournal.Target("WORKBOOK", "Current workbook")),
        ExecutionJournalTargetResolver.resolve(workbookSummary, ExecutionJournalLevel.NORMAL));
    assertEquals(
        List.of(
            new ExecutionJournal.Target("CELL", "Cell Budget!B2"),
            new ExecutionJournal.Target("CELL", "Cell Ops!C3")),
        ExecutionJournalTargetResolver.expandedTargets(
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "B2"),
                    new CellSelector.QualifiedAddress("Ops", "C3")))));
    assertEquals(
        new ExecutionJournal.Target("SHEET", "Sheets [Budget, Ops]"),
        summaryTarget(new SheetSelector.ByNames(List.of("Budget", "Ops"))));
    assertEquals(
        new ExecutionJournal.Target("SHEET", "All sheets"), summaryTarget(new SheetSelector.All()));
    assertEquals(
        new ExecutionJournal.Target("SHEET", "Sheet Budget"),
        summaryTarget(new SheetSelector.ByName("Budget")));
    assertEquals(
        new ExecutionJournal.Target("CELL", "All used cells on Budget"),
        summaryTarget(new CellSelector.AllUsedInSheet("Budget")));
    assertEquals(
        new ExecutionJournal.Target("CELL", "Cell Budget!B2"),
        summaryTarget(new CellSelector.ByAddress("Budget", "B2")));
    assertEquals(
        new ExecutionJournal.Target("CELL", "Cells Budget![B2, C3]"),
        summaryTarget(new CellSelector.ByAddresses("Budget", List.of("B2", "C3"))));
    assertEquals(
        new ExecutionJournal.Target("CELL", "Qualified cells [Budget!B2, Ops!C3]"),
        summaryTarget(
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "B2"),
                    new CellSelector.QualifiedAddress("Ops", "C3")))));
    assertEquals(
        new ExecutionJournal.Target("RANGE", "All ranges on Budget"),
        summaryTarget(new RangeSelector.AllOnSheet("Budget")));
    assertEquals(
        new ExecutionJournal.Target("RANGE", "Range Budget!A1:B4"),
        summaryTarget(new RangeSelector.ByRange("Budget", "A1:B4")));
    assertEquals(
        new ExecutionJournal.Target("RANGE", "Ranges Budget![A1:B4, D2:E3]"),
        summaryTarget(new RangeSelector.ByRanges("Budget", List.of("A1:B4", "D2:E3"))));
    assertEquals(
        new ExecutionJournal.Target("RANGE", "Window Budget!B2:C4"),
        summaryTarget(new RangeSelector.RectangularWindow("Budget", "B2", 3, 2)));
    assertEquals(
        new ExecutionJournal.Target("ROW_BAND", "Rows Budget[1:3]"),
        summaryTarget(new RowBandSelector.Span("Budget", 1, 3)));
    assertEquals(
        new ExecutionJournal.Target("ROW_BAND", "Row insertion Budget[before=2, count=4]"),
        summaryTarget(new RowBandSelector.Insertion("Budget", 2, 4)));
    assertEquals(
        new ExecutionJournal.Target("COLUMN_BAND", "Columns Budget[1:3]"),
        summaryTarget(new ColumnBandSelector.Span("Budget", 1, 3)));
    assertEquals(
        new ExecutionJournal.Target("COLUMN_BAND", "Column insertion Budget[before=2, count=4]"),
        summaryTarget(new ColumnBandSelector.Insertion("Budget", 2, 4)));
    assertEquals(
        new ExecutionJournal.Target("DRAWING", "All drawing objects on Budget"),
        summaryTarget(new DrawingObjectSelector.AllOnSheet("Budget")));
    assertEquals(
        new ExecutionJournal.Target("DRAWING", "Drawing object Budget!Logo"),
        summaryTarget(new DrawingObjectSelector.ByName("Budget", "Logo")));
    assertEquals(
        new ExecutionJournal.Target("CHART", "All charts on Budget"),
        summaryTarget(new ChartSelector.AllOnSheet("Budget")));
    assertEquals(
        new ExecutionJournal.Target("CHART", "Chart Budget!Revenue"),
        summaryTarget(new ChartSelector.ByName("Budget", "Revenue")));
    assertEquals(
        new ExecutionJournal.Target("TABLE", "Tables [BudgetTable, OpsTable]"),
        summaryTarget(new TableSelector.ByNames(List.of("BudgetTable", "OpsTable"))));
    assertEquals(
        new ExecutionJournal.Target("TABLE", "All tables"), summaryTarget(new TableSelector.All()));
    assertEquals(
        new ExecutionJournal.Target("TABLE", "Table BudgetTable"),
        summaryTarget(new TableSelector.ByName("BudgetTable")));
    assertEquals(
        new ExecutionJournal.Target("TABLE", "Table Budget!BudgetTable"),
        summaryTarget(new TableSelector.ByNameOnSheet("BudgetTable", "Budget")));
    assertEquals(
        new ExecutionJournal.Target("PIVOT_TABLE", "Pivot tables [RevenuePivot, CostPivot]"),
        summaryTarget(new PivotTableSelector.ByNames(List.of("RevenuePivot", "CostPivot"))));
    assertEquals(
        new ExecutionJournal.Target("PIVOT_TABLE", "All pivot tables"),
        summaryTarget(new PivotTableSelector.All()));
    assertEquals(
        new ExecutionJournal.Target("PIVOT_TABLE", "Pivot table RevenuePivot"),
        summaryTarget(new PivotTableSelector.ByName("RevenuePivot")));
    assertEquals(
        new ExecutionJournal.Target("PIVOT_TABLE", "Pivot table Budget!RevenuePivot"),
        summaryTarget(new PivotTableSelector.ByNameOnSheet("RevenuePivot", "Budget")));
    assertEquals(
        new ExecutionJournal.Target("NAMED_RANGE", "Named ranges [Gross, Margin]"),
        summaryTarget(new NamedRangeSelector.ByNames(List.of("Gross", "Margin"))));
    assertEquals(
        new ExecutionJournal.Target("NAMED_RANGE", "All named ranges"),
        summaryTarget(new NamedRangeSelector.All()));
    assertEquals(
        new ExecutionJournal.Target("NAMED_RANGE", "Named range Gross"),
        summaryTarget(new NamedRangeSelector.ByName("Gross")));
    assertEquals(
        new ExecutionJournal.Target("NAMED_RANGE", "Workbook-scoped named range Gross"),
        summaryTarget(new NamedRangeSelector.WorkbookScope("Gross")));
    assertEquals(
        new ExecutionJournal.Target("NAMED_RANGE", "Sheet-scoped named range Budget!Margin"),
        summaryTarget(new NamedRangeSelector.SheetScope("Margin", "Budget")));

    ExecutionJournal.Target anyOfSummary =
        summaryTarget(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.ByName("Gross"),
                    new NamedRangeSelector.WorkbookScope("Margin"),
                    new NamedRangeSelector.SheetScope("Net", "Budget"))));
    assertEquals("NAMED_RANGE", anyOfSummary.kind());
    assertTrue(anyOfSummary.label().startsWith("Named range selector ["));

    assertEquals(
        new ExecutionJournal.Target("TABLE_ROW", "All rows in Table BudgetTable"),
        summaryTarget(new TableRowSelector.AllRows(new TableSelector.ByName("BudgetTable"))));
    assertEquals(
        new ExecutionJournal.Target("TABLE_ROW", "Row 3 in Table BudgetTable"),
        summaryTarget(new TableRowSelector.ByIndex(new TableSelector.ByName("BudgetTable"), 3)));
    assertEquals(
        new ExecutionJournal.Target(
            "TABLE_ROW", "Row where Item=Text[text=Hosting] in Table BudgetTable"),
        summaryTarget(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Text(text("Hosting")))));
    assertEquals(
        new ExecutionJournal.Target("TABLE_ROW", "Row where Item=Blank[] in Table BudgetTable"),
        summaryTarget(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"), "Item", new CellInput.Blank())));
    assertEquals(
        new ExecutionJournal.Target(
            "TABLE_ROW", "Row where Item=Number[number=42.0] in Table BudgetTable"),
        summaryTarget(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"), "Item", new CellInput.Numeric(42.0d))));
    assertEquals(
        new ExecutionJournal.Target(
            "TABLE_ROW", "Row where Item=Boolean[value=true] in Table BudgetTable"),
        summaryTarget(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.BooleanValue(true))));
    assertEquals(
        new ExecutionJournal.Target(
            "TABLE_ROW", "Row where Item=Formula[text=A1*2] in Table BudgetTable"),
        summaryTarget(
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByName("BudgetTable"),
                "Item",
                new CellInput.Formula(text("A1*2")))));
    assertEquals(
        new ExecutionJournal.Target("TABLE_CELL", "Column Amount in Row 3 in Table BudgetTable"),
        summaryTarget(
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByIndex(new TableSelector.ByName("BudgetTable"), 3),
                "Amount")));

    List<ExecutionJournal.Target> expandedPivotNames =
        expandedTargets(new PivotTableSelector.ByNames(List.of("RevenuePivot", "CostPivot")));
    assertEquals(
        List.of(
            new ExecutionJournal.Target("PIVOT_TABLE", "Pivot table RevenuePivot"),
            new ExecutionJournal.Target("PIVOT_TABLE", "Pivot table CostPivot")),
        expandedPivotNames);

    List<ExecutionJournal.Target> expandedNamedRangeNames =
        expandedTargets(new NamedRangeSelector.ByNames(List.of("Gross", "Margin")));
    assertEquals(
        List.of(
            new ExecutionJournal.Target("NAMED_RANGE", "Named range Gross"),
            new ExecutionJournal.Target("NAMED_RANGE", "Named range Margin")),
        expandedNamedRangeNames);

    List<ExecutionJournal.Target> expandedAnyOf =
        expandedTargets(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.ByName("Gross"),
                    new NamedRangeSelector.WorkbookScope("Margin"),
                    new NamedRangeSelector.SheetScope("Net", "Budget"))));
    assertEquals(
        List.of(
            new ExecutionJournal.Target("NAMED_RANGE", "Named range Gross"),
            new ExecutionJournal.Target("NAMED_RANGE", "Workbook-scoped named range Margin"),
            new ExecutionJournal.Target("NAMED_RANGE", "Sheet-scoped named range Budget!Net")),
        expandedAnyOf);

    assertEquals(
        List.of(new ExecutionJournal.Target("TABLE_ROW", "Row 3 in Table BudgetTable")),
        expandedTargets(new TableRowSelector.ByIndex(new TableSelector.ByName("BudgetTable"), 3)));
  }

  @Test
  void recorderBuildsVerboseFailureAndGuardsPhaseReuse() {
    WorkbookPlan request = verbosePlan();
    List<ExecutionJournal.Event> emitted = new ArrayList<>();
    ExecutionJournalRecorder recorder = ExecutionJournalRecorder.start(request, emitted::add);

    recorder.setWarnings(null);
    ExecutionJournalRecorder.PhaseHandle validation = recorder.beginValidation();
    validation.succeed();
    assertThrows(IllegalStateException.class, validation::succeed);

    InspectionStep successStep =
        new InspectionStep(
            "step-1", new WorkbookSelector.Current(), new InspectionQuery.GetWorkbookSummary());
    recorder.beginStep(0, successStep).succeed();

    InspectionStep failedWithoutCalculation =
        new InspectionStep(
            "step-2", new TableSelector.ByName("BudgetTable"), new InspectionQuery.GetTables());
    recorder
        .beginStep(1, failedWithoutCalculation)
        .fail(
            GridGrindProblemCode.INVALID_REQUEST,
            GridGrindProblemCategory.REQUEST,
            "EXECUTE_STEP",
            "bad table");

    InspectionStep failedWithCalculation =
        new InspectionStep(
            "step-3", new CellSelector.ByAddress("Budget", "B2"), new InspectionQuery.GetCells());
    ExecutionJournalRecorder.StepHandle stepHandle = recorder.beginStep(2, failedWithCalculation);
    ExecutionJournalRecorder.PhaseHandle calculationPreflight =
        recorder.beginCalculationPreflight();
    calculationPreflight.succeed();
    ExecutionJournalRecorder.PhaseHandle calculationExecution =
        recorder.beginCalculationExecution();
    calculationExecution.fail("failed (IO_ERROR)");
    stepHandle.fail(
        GridGrindProblemCode.IO_ERROR, GridGrindProblemCategory.IO, "EXECUTE_STEP", "disk issue");

    ExecutionJournal journal = recorder.buildFailure(3, GridGrindProblemCode.IO_ERROR, 2, "step-3");

    assertEquals(List.of(), journal.warnings());
    assertFalse(journal.events().isEmpty());
    assertFalse(emitted.isEmpty());
    assertEquals(ExecutionJournal.Status.SUCCEEDED, journal.calculation().preflight().status());
    assertEquals(ExecutionJournal.Status.FAILED, journal.calculation().execution().status());
    assertEquals(ExecutionJournal.Status.FAILED, journal.outcome().status());
    assertEquals(GridGrindProblemCode.IO_ERROR, journal.outcome().failureCode().orElseThrow());
  }

  @Test
  void recorderSupportsNullRequestAndCalculationIoFailures() throws Exception {
    ExecutionJournalRecorder recorder =
        ExecutionJournalRecorder.start(null, ExecutionJournalSink.NOOP);
    ExecutionJournal unknownJournal = recorder.buildSuccess(0);
    assertEquals(java.util.Optional.empty(), unknownJournal.planId());
    assertEquals(java.util.Optional.empty(), unknownJournal.source().type());
    assertEquals(java.util.Optional.empty(), unknownJournal.persistence().type());

    WorkbookPlan request = verbosePlan();
    ExecutionJournalRecorder verboseRecorder =
        ExecutionJournalRecorder.start(request, ExecutionJournalSink.NOOP);
    ExecutionJournalRecorder.PhaseHandle calculationPreflight =
        verboseRecorder.beginCalculationPreflight();
    calculationPreflight.succeed();
    ExecutionJournalRecorder.PhaseHandle calculationExecution =
        verboseRecorder.beginCalculationExecution();
    calculationExecution.fail("failed (IO_ERROR)");
    ExecutionJournalRecorder.StepHandle stepHandle =
        verboseRecorder.beginStep(
            0,
            new InspectionStep(
                "step-1",
                new CellSelector.ByAddress("Budget", "B2"),
                new InspectionQuery.GetCells()));
    stepHandle.fail(
        GridGrindProblemCode.IO_ERROR, GridGrindProblemCategory.IO, "EXECUTE_STEP", "boom");
    ExecutionJournal journal =
        verboseRecorder.buildFailure(1, GridGrindProblemCode.IO_ERROR, 0, "step-1");
    assertEquals(ExecutionJournal.Status.FAILED, journal.calculation().execution().status());
  }

  private static WorkbookPlan verbosePlan() {
    return new WorkbookPlan(
        GridGrindProtocolVersion.current(),
        "phase-5-plan",
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        new ExecutionPolicyInput(
            ExecutionModeInput.defaults(),
            new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
        FormulaEnvironmentInput.empty(),
        List.of());
  }

  private static List<ExecutionJournal.Target> expandedTargets(Selector selector) {
    return ExecutionJournalTargetResolver.expandedTargets(selector);
  }

  private static ExecutionJournal.Target summaryTarget(Selector selector) {
    return ExecutionJournalTargetResolver.summaryTarget(selector);
  }

  private static dev.erst.gridgrind.contract.source.TextSourceInput text(String value) {
    return dev.erst.gridgrind.contract.source.TextSourceInput.inline(value);
  }
}
