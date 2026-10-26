package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.WorkbookCellCommand;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookSheetCommand;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Verifies formula-origin tracking without evaluating ordinary or opaque formula cells. */
class FormulaOriginTrackerTest {
  @Test
  void plansEveryCoordinateForOffsetSetAndClearRanges() throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      FormulaOriginTracker.FormulaWrites setWrites =
          tracker.plannedWrites(
              workbook,
              new WorkbookCellCommand.SetRange(
                  "Ops",
                  "B2:C3",
                  List.of(
                      List.of(new ResolvedFormula("1+1"), new CellValue.Text("a")),
                      List.of(new CellValue.Text("b"), new ResolvedFormula("2+2")))));
      FormulaOriginTracker.FormulaWrites clearWrites =
          tracker.plannedWrites(workbook, new WorkbookCellCommand.ClearRange("Ops", "B2:C3"));
      FormulaOriginTracker.FormulaWrites clearArrayWrites =
          tracker.plannedWrites(workbook, new WorkbookCellCommand.ClearArrayFormula("Ops", "D5"));

      Set<FormulaOriginTracker.CellCoordinate> expected =
          Set.of(
              new FormulaOriginTracker.CellCoordinate("Ops", "B2"),
              new FormulaOriginTracker.CellCoordinate("Ops", "C2"),
              new FormulaOriginTracker.CellCoordinate("Ops", "B3"),
              new FormulaOriginTracker.CellCoordinate("Ops", "C3"));
      assertEquals(expected, setWrites.cells().keySet());
      assertEquals(expected, clearWrites.cells().keySet());
      assertEquals(
          Set.of(new FormulaOriginTracker.CellCoordinate("Ops", "D5")),
          clearArrayWrites.cells().keySet());
      assertTrue(clearArrayWrites.cells().values().stream().allMatch(Optional::isEmpty));
    }
  }

  @Test
  void tracksEveryNormalFormulaAuthoringFamilyAndExcludesRawFormulas() throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    WorkbookExecutionEngine engine = new WorkbookExecutionEngine();
    StepReference author = new StepReference(3, new StepId("author"), "MUTATION", "SET_CELL");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      engine.apply(workbook, new WorkbookSheetCommand.CreateSheet("Ops"));
      workbook.xssfWorkbook().getSheet("Ops").createRow(7).createCell(0).setBlank();

      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A1", new ResolvedFormula("1+1")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "B1", new ResolvedFormula("2+2")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetRange(
              "Ops",
              "B1:C1",
              List.of(List.of(new ResolvedFormula("2+2"), new CellValue.Text("plain")))),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.AppendRow(
              "Ops", List.of(new ResolvedFormula("3+3"), new CellValue.Text("tail"))),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetArrayFormula("Ops", "A3:A4", new ResolvedArrayFormula("4+4")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "D1", new ResolvedRawFormula("RAW()")),
          author);

      assertEquals(author, tracker.originFor("Ops", "A1").orElseThrow());
      assertEquals(author, tracker.originFor("Ops", "B1").orElseThrow());
      assertEquals(author, tracker.originFor("Ops", "A2").orElseThrow());
      assertEquals(author, tracker.originFor("Ops", "A3").orElseThrow());
      assertEquals(author, tracker.originFor("Ops", "A4").orElseThrow());
      assertTrue(tracker.originFor("Ops", "C1").isEmpty());
      assertTrue(tracker.originFor("Ops", "D1").isEmpty());
      assertEquals(
          author,
          tracker
              .originFor(new InvalidFormulaException("Ops", "A1", "1+1", "bad", null))
              .orElseThrow());
      assertTrue(
          tracker.originFor(new IllegalArgumentException("not a formula failure")).isEmpty());
      assertTrue(tracker.originFor(new AssertionError("not an exception")).isEmpty());
    }
  }

  @Test
  void removesOriginsWhenLaterMutationsReplaceOrClearFormulaCells() throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    WorkbookExecutionEngine engine = new WorkbookExecutionEngine();
    StepReference author = new StepReference(0, new StepId("author"), "MUTATION", "SET_CELL");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      engine.apply(workbook, new WorkbookSheetCommand.CreateSheet("Ops"));
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A1", new ResolvedFormula("1+1")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "B1", new ResolvedFormula("2+2")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A1", new CellValue.Text("replaced")),
          new StepReference(1, new StepId("replace"), "MUTATION", "SET_CELL"));
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A1", new ResolvedFormula("2+2")),
          new StepReference(2, new StepId("rewrite"), "MUTATION", "SET_CELL"));
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.ClearRange("Ops", "A1:A1"),
          new StepReference(3, new StepId("clear"), "MUTATION", "CLEAR_RANGE"));
      assertTrue(tracker.originFor("Ops", "A1").isEmpty());
      assertTrue(
          tracker
              .plannedWrites(workbook, new WorkbookSheetCommand.SetActiveSheet("Ops"))
              .cells()
              .isEmpty());
    }
  }

  @Test
  void calculationFailuresRetainFormulaAuthorsWhenAnExecutionExceptionSurfacesLater()
      throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    WorkbookExecutionEngine engine = new WorkbookExecutionEngine();
    StepReference author = new StepReference(1, new StepId("author"), "MUTATION", "SET_CELL");
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      engine.apply(workbook, new WorkbookSheetCommand.CreateSheet("Ops"));
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A1", new ResolvedFormula("1+1")),
          author);

      dev.erst.gridgrind.core.problem.GridGrindProblemDetail.Problem problem =
          new ExecutionCalculationSupport(ignored -> {})
              .calculationProblemFor(
                  request,
                  new CalculationPolicyExecutor.FailureDetail(
                      GridGrindProblemCode.INVALID_FORMULA,
                      CalculationPolicyExecutor.Phase.EXECUTION,
                      "Ops",
                      "A1",
                      "1+1",
                      "invalid formula",
                      new InvalidFormulaException("Ops", "A1", "1+1", "invalid formula", null)),
                  tracker,
                  author);

      ProblemContext.ExecuteStep context =
          assertInstanceOf(ProblemContext.ExecuteStep.class, problem.context());
      assertEquals(author, context.step());
      assertTrue(context.surfacedAtStep().isEmpty());

      assertInstanceOf(
          ProblemContext.ExecuteCalculation.Execution.class,
          new ExecutionCalculationSupport(ignored -> {})
              .calculationProblemFor(
                  request,
                  new CalculationPolicyExecutor.FailureDetail(
                      GridGrindProblemCode.INVALID_FORMULA,
                      CalculationPolicyExecutor.Phase.EXECUTION,
                      null,
                      null,
                      null,
                      "missing formula location",
                      null),
                  tracker,
                  author)
              .context());
      assertInstanceOf(
          ProblemContext.ExecuteCalculation.Execution.class,
          new ExecutionCalculationSupport(ignored -> {})
              .calculationProblemFor(
                  request,
                  new CalculationPolicyExecutor.FailureDetail(
                      GridGrindProblemCode.INVALID_FORMULA,
                      CalculationPolicyExecutor.Phase.EXECUTION,
                      "Ops",
                      null,
                      null,
                      "missing formula address",
                      null),
                  tracker,
                  author)
              .context());
    }
  }

  @Test
  void prunesEveryChangedFormulaTopologyBeforeRecordingTheNextMutation() throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    WorkbookExecutionEngine engine = new WorkbookExecutionEngine();
    StepReference author = new StepReference(0, new StepId("author"), "MUTATION", "SET_CELL");
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      engine.apply(workbook, new WorkbookSheetCommand.CreateSheet("Ops"));
      engine.apply(workbook, new WorkbookSheetCommand.CreateSheet("Removed"));
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetRange(
              "Ops",
              "A1:C1",
              List.of(
                  List.of(
                      new ResolvedFormula("1+1"),
                      new ResolvedFormula("2+2"),
                      new ResolvedFormula("3+3")))),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A5", new ResolvedFormula("5+5")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Ops", "A6", new ResolvedFormula("6+6")),
          author);
      applyAndRecord(
          tracker,
          engine,
          workbook,
          new WorkbookCellCommand.SetCell("Removed", "A1", new ResolvedFormula("7+7")),
          author);

      assertEquals(author, tracker.originFor("Ops", "A1").orElseThrow());
      workbook.xssfWorkbook().getSheet("Ops").getRow(0).getCell(1).setBlank();
      workbook.xssfWorkbook().getSheet("Ops").getRow(0).getCell(2).setCellFormula("9+9");
      workbook
          .xssfWorkbook()
          .getSheet("Ops")
          .removeRow(workbook.xssfWorkbook().getSheet("Ops").getRow(4));
      workbook
          .xssfWorkbook()
          .getSheet("Ops")
          .getRow(5)
          .removeCell(workbook.xssfWorkbook().getSheet("Ops").getRow(5).getCell(0));
      workbook.xssfWorkbook().removeSheetAt(workbook.xssfWorkbook().getSheetIndex("Removed"));

      tracker.record(workbook, FormulaOriginTracker.FormulaWrites.none(), author);

      assertEquals(author, tracker.originFor("Ops", "A1").orElseThrow());
      assertTrue(tracker.originFor("Ops", "B1").isEmpty());
      assertTrue(tracker.originFor("Ops", "C1").isEmpty());
      assertTrue(tracker.originFor("Ops", "A5").isEmpty());
      assertTrue(tracker.originFor("Ops", "A6").isEmpty());
      assertTrue(tracker.originFor("Removed", "A1").isEmpty());
    }
  }

  private static void applyAndRecord(
      FormulaOriginTracker tracker,
      WorkbookExecutionEngine engine,
      ExcelWorkbook workbook,
      WorkbookCommand command,
      StepReference author) {
    FormulaOriginTracker.FormulaWrites writes = tracker.plannedWrites(workbook, command);
    engine.apply(workbook, command);
    tracker.record(workbook, writes, author);
  }
}
