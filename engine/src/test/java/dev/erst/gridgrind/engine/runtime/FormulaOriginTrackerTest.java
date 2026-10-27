package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.ResolvedArrayFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.ResolvedMutationExecutor;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.formulas.InvalidFormulaException;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import org.junit.jupiter.api.Test;

/** Verifies formula-origin tracking without evaluating ordinary or opaque formula cells. */
class FormulaOriginTrackerTest {
  @Test
  void plansEveryCoordinateForOffsetSetAndClearRanges() throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      FormulaOriginTracker.FormulaWrites setWrites =
          tracker.plannedWrites(
              workbook,
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("B2:C3")),
                  new ResolvedMutationAction.SetRange(
                      List.of(
                          List.of(new ResolvedFormula("1+1"), new CellValue.Text("a")),
                          List.of(new CellValue.Text("b"), new ResolvedFormula("2+2"))))));
      FormulaOriginTracker.FormulaWrites clearWrites =
          tracker.plannedWrites(
              workbook,
              new ResolvedStep(
                  new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("B2:C3")),
                  new CellMutationAction.ClearRange()));
      FormulaOriginTracker.FormulaWrites clearArrayWrites =
          tracker.plannedWrites(
              workbook,
              new ResolvedStep(
                  new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("D5")),
                  new CellMutationAction.ClearArrayFormula()));

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
    StepReference author = new StepReference(3, new StepId("author"), "MUTATION", "SET_CELL");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      workbook.xssfWorkbook().getSheet("Ops").createRow(7).createCell(0).setBlank();

      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("1+1"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("B1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("2+2"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("B1:C1")),
              new ResolvedMutationAction.SetRange(
                  List.of(List.of(new ResolvedFormula("2+2"), new CellValue.Text("plain"))))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new ResolvedMutationAction.AppendRow(
                  List.of(new ResolvedFormula("3+3"), new CellValue.Text("tail")))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("A3:A4")),
              new ResolvedMutationAction.SetArrayFormula(new ResolvedArrayFormula("4+4"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("D1")),
              new ResolvedMutationAction.SetCell(new ResolvedRawFormula("RAW()"))),
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
    StepReference author = new StepReference(0, new StepId("author"), "MUTATION", "SET_CELL");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("1+1"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("B1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("2+2"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A1")),
              new ResolvedMutationAction.SetCell(new CellValue.Text("replaced"))),
          new StepReference(1, new StepId("replace"), "MUTATION", "SET_CELL"));
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("2+2"))),
          new StepReference(2, new StepId("rewrite"), "MUTATION", "SET_CELL"));
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("A1:A1")),
              new CellMutationAction.ClearRange()),
          new StepReference(3, new StepId("clear"), "MUTATION", "CLEAR_RANGE"));
      assertTrue(tracker.originFor("Ops", "A1").isEmpty());
      assertTrue(
          tracker
              .plannedWrites(
                  workbook,
                  new ResolvedStep(
                      new SheetSelector.ByName(new SheetName("Ops")),
                      new WorkbookMutationAction.SetActiveSheet()))
              .cells()
              .isEmpty());
    }
  }

  @Test
  void calculationFailuresRetainFormulaAuthorsWhenAnExecutionExceptionSurfacesLater()
      throws Exception {
    FormulaOriginTracker tracker = new FormulaOriginTracker();
    StepReference author = new StepReference(1, new StepId("author"), "MUTATION", "SET_CELL");
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("1+1"))),
          author);

      dev.erst.gridgrind.core.problem.GridGrindProblemDetail.Problem problem =
          new CalculationWorkflow(ignored -> {})
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
          new CalculationWorkflow(ignored -> {})
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
          new CalculationWorkflow(ignored -> {})
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
    StepReference author = new StepReference(0, new StepId("author"), "MUTATION", "SET_CELL");
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Ops")),
              new WorkbookMutationAction.EnsureSheet()));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new SheetSelector.ByName(new SheetName("Removed")),
              new WorkbookMutationAction.EnsureSheet()));
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new RangeSelector.ByRange(new SheetName("Ops"), new RangeRef("A1:C1")),
              new ResolvedMutationAction.SetRange(
                  List.of(
                      List.of(
                          new ResolvedFormula("1+1"),
                          new ResolvedFormula("2+2"),
                          new ResolvedFormula("3+3"))))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A5")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("5+5"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Ops"), new CellAddress("A6")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("6+6"))),
          author);
      applyAndRecord(
          tracker,
          workbook,
          new ResolvedStep(
              new CellSelector.ByAddress(new SheetName("Removed"), new CellAddress("A1")),
              new ResolvedMutationAction.SetCell(new ResolvedFormula("7+7"))),
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
      XssfWorkbook workbook,
      ResolvedStep command,
      StepReference author) {
    FormulaOriginTracker.FormulaWrites writes = tracker.plannedWrites(workbook, command);
    ResolvedMutationExecutor.apply(workbook, command);
    tracker.record(workbook, writes, author);
  }
}
