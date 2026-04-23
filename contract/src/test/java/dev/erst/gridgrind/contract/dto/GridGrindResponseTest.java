package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for step-based successful and failed protocol responses. */
class GridGrindResponseTest {
  @Test
  void successDefaultsPersistenceAndCopiesWarningsAndInspections() {
    List<RequestWarning> warnings = new ArrayList<>();
    warnings.add(
        new RequestWarning(0, "set-total", "SET_CELL", "Quote spaced sheet names in formulas."));
    List<InspectionResult> inspections = new ArrayList<>();
    inspections.add(
        new InspectionResult.WorkbookSummaryResult(
            "summary",
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), 0, false)));

    GridGrindResponse.Success success =
        new GridGrindResponse.Success(null, null, warnings, List.of(), inspections);

    warnings.clear();
    inspections.clear();

    assertEquals(GridGrindProtocolVersion.current(), success.protocolVersion());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, success.persistence());
    assertNull(success.journal().planId());
    assertNull(success.journal().source().type());
    assertNull(success.journal().persistence().type());
    assertEquals(ExecutionJournal.Status.SUCCEEDED, success.journal().outcome().status());
    assertEquals(1, success.warnings().size());
    assertEquals(1, success.inspections().size());
  }

  @Test
  void failureBackfillConstructorCreatesFailedSyntheticJournal() {
    GridGrindResponse.Failure failure =
        new GridGrindResponse.Failure(
            null,
            GridGrindResponse.Problem.of(
                GridGrindProblemCode.INTERNAL_ERROR,
                "boom",
                new GridGrindResponse.ProblemContext.ExecuteRequest(null, null)));

    assertNull(failure.journal().planId());
    assertNull(failure.journal().source().type());
    assertNull(failure.journal().persistence().type());
    assertEquals(ExecutionJournal.Status.FAILED, failure.journal().outcome().status());
  }

  @Test
  void requestWarningsRequireStepIdentity() {
    assertThrows(
        IllegalArgumentException.class, () -> new RequestWarning(-1, "a", "SET_CELL", "warn"));
    assertThrows(
        IllegalArgumentException.class, () -> new RequestWarning(0, " ", "SET_CELL", "warn"));
    assertThrows(IllegalArgumentException.class, () -> new RequestWarning(0, "a", " ", "warn"));
    assertThrows(IllegalArgumentException.class, () -> new RequestWarning(0, "a", "SET_CELL", " "));
  }

  @Test
  void executeStepContextMergesExceptionDataWithoutOverwritingExistingValues() {
    GridGrindResponse.ProblemContext.ExecuteStep base =
        new GridGrindResponse.ProblemContext.ExecuteStep(
            "EXISTING",
            "SAVE_AS",
            2,
            "formula-health",
            "INSPECTION",
            "ANALYZE_FORMULA_HEALTH",
            "Summary",
            null,
            null,
            null,
            null);

    GridGrindResponse.ProblemContext.ExecuteStep enriched =
        base.withExceptionData("Ignored", "B4", "B4:B9", "SUM(B2:B3)", "BudgetTotal");

    assertEquals("Summary", enriched.sheetName());
    assertEquals("B4", enriched.address());
    assertEquals("B4:B9", enriched.range());
    assertEquals("SUM(B2:B3)", enriched.formula());
    assertEquals("BudgetTotal", enriched.namedRangeName());
  }

  @Test
  void parseArgumentsAndProblemsExposeStepCentricContext() {
    GridGrindResponse.ProblemContext.ParseArguments parseArguments =
        new GridGrindResponse.ProblemContext.ParseArguments("--request");
    GridGrindResponse.Problem problem =
        GridGrindResponse.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown field 'reads'",
            new GridGrindResponse.ProblemContext.ExecuteStep(
                "NEW",
                "NONE",
                1,
                "cells",
                "INSPECTION",
                "GET_CELLS",
                "Budget",
                "A1",
                null,
                null,
                null));

    assertEquals("PARSE_ARGUMENTS", parseArguments.stage());
    assertEquals("--request", parseArguments.argument());
    assertEquals("EXECUTE_STEP", problem.context().stage());
    assertEquals(1, problem.context().stepIndex());
    assertEquals("cells", problem.context().stepId());
    assertEquals("GET_CELLS", problem.context().stepType());
    assertEquals("Budget", problem.context().sheetName());
    assertEquals("A1", problem.context().address());
    assertNull(problem.context().range());
    assertTrue(problem.title().contains("request"));
  }

  @Test
  void defaultInterfaceAccessorsReturnNullAndContextMergersPreserveExistingValues() {
    GridGrindResponse.NamedRangeReport formulaOnly =
        new GridGrindResponse.NamedRangeReport.FormulaReport(
            "BudgetExpr", new NamedRangeScope.Workbook(), "SUM(Budget!A1:A3)");
    GridGrindResponse.SheetProtectionReport unprotected =
        new GridGrindResponse.SheetProtectionReport.Unprotected();
    GridGrindResponse.CellReport blankCell =
        new GridGrindResponse.CellReport.BlankReport("A1", "BLANK", "", minimalStyle(), null, null);
    GridGrindResponse.ProblemContext.ReadRequest readRequest =
        new GridGrindResponse.ProblemContext.ReadRequest("/tmp/request.json", "steps[0]", 4, 12);

    assertNull(formulaOnly.target());
    assertNull(unprotected.settings());
    assertNull(blankCell.formula());
    assertNull(blankCell.stringValue());
    assertNull(blankCell.richText());
    assertNull(blankCell.numberValue());
    assertNull(blankCell.booleanValue());
    assertNull(blankCell.errorValue());
    assertNull(readRequest.sourceType());
    assertNull(readRequest.persistenceType());
    assertNull(readRequest.responsePath());
    assertNull(readRequest.sourceWorkbookPath());
    assertNull(readRequest.persistencePath());
    assertNull(readRequest.stepIndex());
    assertNull(readRequest.stepId());
    assertNull(readRequest.stepKind());
    assertNull(readRequest.stepType());
    assertNull(readRequest.sheetName());
    assertNull(readRequest.address());
    assertNull(readRequest.range());
    assertNull(readRequest.formula());
    assertNull(readRequest.namedRangeName());
    assertNull(readRequest.argument());

    GridGrindResponse.ProblemContext.ReadRequest mergedRead =
        readRequest.withJson("ignored", 9, 22);
    assertEquals("steps[0]", mergedRead.jsonPath());
    assertEquals(4, mergedRead.jsonLine());
    assertEquals(12, mergedRead.jsonColumn());

    GridGrindResponse.ProblemContext.ExecuteStep mergedExecute =
        new GridGrindResponse.ProblemContext.ExecuteStep(
                "NEW", "NONE", 1, "cells", "INSPECTION", "GET_CELLS", null, null, null, null, null)
            .withExceptionData("Budget", "A1", "A1:B2", "SUM(A1)", "BudgetTotal");
    assertEquals("Budget", mergedExecute.sheetName());
    assertEquals("A1", mergedExecute.address());
    assertEquals("A1:B2", mergedExecute.range());
    assertEquals("SUM(A1)", mergedExecute.formula());
    assertEquals("BudgetTotal", mergedExecute.namedRangeName());
  }

  @Test
  void validatesPersistenceOutcomesWorkbookSummariesAndProblemCauses() {
    assertEquals(
        "requestedPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.PersistenceOutcome.SavedAs(" ", "/tmp/out.xlsx"))
            .getMessage());
    assertEquals(
        "executionPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.PersistenceOutcome.Overwritten("budget.xlsx", " "))
            .getMessage());
    assertEquals(
        "sheetCount must be 0 for an empty workbook",
        assertThrows(
                IllegalArgumentException.class,
                () -> new GridGrindResponse.WorkbookSummary.Empty(1, List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "selectedSheetNames must only contain values present in sheetNames",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Ops"), 0, false))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.ProblemCause(
                        GridGrindProblemCode.INVALID_REQUEST, " ", "READ_REQUEST"))
            .getMessage());
    assertEquals(
        "stage must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindResponse.ProblemCause(
                        GridGrindProblemCode.INVALID_REQUEST, "bad request", " "))
            .getMessage());
    assertEquals(
        List.of(),
        new GridGrindResponse.Problem(
                GridGrindProblemCode.INVALID_REQUEST,
                GridGrindProblemCode.INVALID_REQUEST.category(),
                GridGrindProblemCode.INVALID_REQUEST.recovery(),
                GridGrindProblemCode.INVALID_REQUEST.title(),
                "bad request",
                GridGrindProblemCode.INVALID_REQUEST.resolution(),
                new GridGrindResponse.ProblemContext.ValidateRequest("NEW", "NONE"),
                null,
                null)
            .causes());
  }

  @Test
  void successCopiesAssertionsAndProblemCanCarryAssertionFailure() {
    List<AssertionResult> assertions = new ArrayList<>();
    assertions.add(new AssertionResult("assert-total", "EXPECT_CELL_VALUE"));

    GridGrindResponse.Success success =
        new GridGrindResponse.Success(
            null,
            null,
            List.of(),
            assertions,
            List.of(
                new InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 0, false))));
    assertions.clear();

    assertEquals(1, success.assertions().size());
    assertEquals("assert-total", success.assertions().getFirst().stepId());

    var failure =
        new dev.erst.gridgrind.contract.assertion.AssertionFailure(
            "assert-total",
            "EXPECT_CELL_VALUE",
            new dev.erst.gridgrind.contract.selector.CellSelector.ByAddress("Budget", "B4"),
            new dev.erst.gridgrind.contract.assertion.Assertion.CellValue(
                new dev.erst.gridgrind.contract.assertion.ExpectedCellValue.NumericValue(42.0d)),
            List.of());
    GridGrindResponse.Problem problem =
        new GridGrindResponse.Problem(
            GridGrindProblemCode.ASSERTION_FAILED,
            GridGrindProblemCode.ASSERTION_FAILED.category(),
            GridGrindProblemCode.ASSERTION_FAILED.recovery(),
            GridGrindProblemCode.ASSERTION_FAILED.title(),
            "assertion failed",
            GridGrindProblemCode.ASSERTION_FAILED.resolution(),
            new GridGrindResponse.ProblemContext.ExecuteStep(
                "NEW",
                "NONE",
                0,
                "assert-total",
                "ASSERTION",
                "EXPECT_CELL_VALUE",
                "Budget",
                "B4",
                null,
                null,
                null),
            failure,
            List.of());
    assertEquals(failure, problem.assertionFailure());
  }

  @Test
  void executionJournalValidatesPhasesAndFailureClassification() {
    ExecutionJournal.Phase successPhase =
        new ExecutionJournal.Phase(
            ExecutionJournal.Status.SUCCEEDED, "2026-04-18T10:00:00Z", "2026-04-18T10:00:01Z", 1);
    ExecutionJournal journal =
        new ExecutionJournal(
            "budget-audit",
            ExecutionJournalLevel.VERBOSE,
            new ExecutionJournal.SourceSummary("NEW", null),
            new ExecutionJournal.PersistenceSummary("SAVE_AS", "/tmp/report.xlsx"),
            successPhase,
            successPhase,
            successPhase,
            new ExecutionJournal.Calculation(successPhase, successPhase),
            successPhase,
            successPhase,
            List.of(
                new ExecutionJournal.Step(
                    0,
                    "assert-total",
                    "ASSERTION",
                    "EXPECT_CELL_VALUE",
                    List.of(new ExecutionJournal.Target("CELL", "Cell Budget!B4")),
                    successPhase,
                    ExecutionJournal.StepOutcome.FAILED,
                    new ExecutionJournal.FailureClassification(
                        GridGrindProblemCode.ASSERTION_FAILED,
                        GridGrindProblemCategory.REQUEST,
                        "EXECUTE_STEP",
                        "observed value mismatch"))),
            List.of(),
            new ExecutionJournal.Outcome(
                ExecutionJournal.Status.FAILED,
                1,
                0,
                22,
                0,
                "assert-total",
                GridGrindProblemCode.ASSERTION_FAILED),
            List.of(
                new ExecutionJournal.Event(
                    "2026-04-18T10:00:00Z", "STEP", "started", 0, "assert-total")));

    assertEquals("budget-audit", journal.planId());
    assertEquals(ExecutionJournalLevel.VERBOSE, journal.level());
    assertEquals("Cell Budget!B4", journal.steps().getFirst().resolvedTargets().getFirst().label());
    assertEquals(
        "NOT_STARTED phases must omit timestamps and use durationMillis=0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.NOT_STARTED, "2026-04-18T10:00:00Z", null, 0))
            .getMessage());
  }

  private static GridGrindResponse.CellStyleReport minimalStyle() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            null,
            false,
            false),
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }
}
