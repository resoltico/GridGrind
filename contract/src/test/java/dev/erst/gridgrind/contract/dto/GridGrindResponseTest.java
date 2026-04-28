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
        GridGrindResponses.success(warnings, List.of(), inspections);
    GridGrindResponse.Success successWithoutInspections =
        GridGrindResponses.success(warnings, List.of(), List.of());

    warnings.clear();
    inspections.clear();

    assertEquals(GridGrindProtocolVersion.current(), success.protocolVersion());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, success.persistence());
    assertEquals(java.util.Optional.empty(), success.journal().planId());
    assertEquals(java.util.Optional.empty(), success.journal().source().type());
    assertEquals(java.util.Optional.empty(), success.journal().persistence().type());
    assertEquals(ExecutionJournal.Status.SUCCEEDED, success.journal().outcome().status());
    assertEquals(1, success.warnings().size());
    assertEquals(1, success.inspections().size());
    assertEquals(List.of(), successWithoutInspections.inspections());
  }

  @Test
  void failureBackfillConstructorCreatesFailedSyntheticJournal() {
    GridGrindResponse.Failure failure =
        GridGrindResponses.failure(
            null,
            GridGrindResponse.Problem.of(
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "boom",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.unknown())));

    assertEquals(java.util.Optional.empty(), failure.journal().planId());
    assertEquals(java.util.Optional.empty(), failure.journal().source().type());
    assertEquals(java.util.Optional.empty(), failure.journal().persistence().type());
    assertEquals(ExecutionJournal.Status.FAILED, failure.journal().outcome().status());
    assertEquals(
        GridGrindProblemCode.INVALID_ARGUMENTS,
        failure.journal().outcome().failureCode().orElseThrow());
  }

  @Test
  void syntheticJournalRejectsInvalidFailureCodeCombinations() {
    IllegalArgumentException missingFailureCode =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                GridGrindResponse.syntheticJournal(
                    ExecutionJournal.Status.FAILED, java.util.Optional.empty()));
    assertEquals(
        "failureCode must be present when status is FAILED", missingFailureCode.getMessage());

    IllegalArgumentException unexpectedFailureCode =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                GridGrindResponse.syntheticJournal(
                    ExecutionJournal.Status.SUCCEEDED,
                    java.util.Optional.of(GridGrindProblemCode.INVALID_ARGUMENTS)));
    assertEquals(
        "failureCode is only permitted when status is FAILED", unexpectedFailureCode.getMessage());
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
  void executeStepContextMergesTypedLocationsWithoutReintroducingNullPadding() {
    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep base =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
            dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known(
                "EXISTING", "SAVE_AS"),
            new dev.erst.gridgrind.contract.dto.ProblemContext.StepReference(
                2, "formula-health", "INSPECTION", "ANALYZE_FORMULA_HEALTH"),
            dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.sheet("Summary"));

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep enriched =
        base.withLocation(
            dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.formulaCell(
                "Ignored", "B4", "SUM(B2:B3)"));

    assertEquals(java.util.Optional.of("Summary"), enriched.sheetName());
    assertEquals(java.util.Optional.of("B4"), enriched.address());
    assertEquals(java.util.Optional.empty(), enriched.range());
    assertEquals(java.util.Optional.of("SUM(B2:B3)"), enriched.formula());
    assertEquals(java.util.Optional.empty(), enriched.namedRangeName());
  }

  @Test
  void parseArgumentsAndProblemsExposeStepCentricContext() {
    dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments parseArguments =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
            dev.erst.gridgrind.contract.dto.ProblemContext.CliArgument.named("--request"));
    GridGrindResponse.Problem problem =
        GridGrindResponse.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown field 'reads'",
            new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known("NEW", "NONE"),
                new dev.erst.gridgrind.contract.dto.ProblemContext.StepReference(
                    1, "cells", "INSPECTION", "GET_CELLS"),
                dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.cell(
                    "Budget", "A1")));
    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep executeStep =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep.class, problem.context());

    assertEquals("PARSE_ARGUMENTS", parseArguments.stage());
    assertEquals(java.util.Optional.of("--request"), parseArguments.argumentName());
    assertEquals("EXECUTE_STEP", executeStep.stage());
    assertEquals(1, executeStep.stepIndex());
    assertEquals("cells", executeStep.stepId());
    assertEquals("GET_CELLS", executeStep.stepType());
    assertEquals(java.util.Optional.of("Budget"), executeStep.sheetName());
    assertEquals(java.util.Optional.of("A1"), executeStep.address());
    assertEquals(java.util.Optional.empty(), executeStep.range());
    assertTrue(problem.title().contains("request"));
  }

  @Test
  void typedVariantsReplaceNullPaddingAndContextMergersPreserveExistingValues() {
    GridGrindResponse.NamedRangeReport.FormulaReport formulaOnly =
        new GridGrindResponse.NamedRangeReport.FormulaReport(
            "BudgetExpr", new NamedRangeScope.Workbook(), "SUM(Budget!A1:A3)");
    GridGrindResponse.SheetProtectionReport.Unprotected unprotected =
        new GridGrindResponse.SheetProtectionReport.Unprotected();
    dev.erst.gridgrind.contract.dto.CellReport.BlankReport blankCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BlankReport(
            "A1",
            "BLANK",
            "",
            minimalStyle(),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest readRequest =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
            dev.erst.gridgrind.contract.dto.ProblemContext.RequestInput.requestFile(
                "/tmp/request.json"),
            dev.erst.gridgrind.contract.dto.ProblemContext.JsonLocation.located("steps[0]", 4, 12));

    assertEquals("BudgetExpr", formulaOnly.name());
    assertEquals("SUM(Budget!A1:A3)", formulaOnly.refersToFormula());
    assertInstanceOf(GridGrindResponse.SheetProtectionReport.Unprotected.class, unprotected);
    assertEquals("BLANK", blankCell.effectiveType());
    assertEquals(java.util.Optional.empty(), blankCell.hyperlink());
    assertEquals(java.util.Optional.empty(), blankCell.comment());
    assertNull(new AutofilterEntryReport.SheetOwned("A1:B2").sortState());
    assertEquals(java.util.Optional.of("/tmp/request.json"), readRequest.requestPath());
    assertEquals(java.util.Optional.of("steps[0]"), readRequest.jsonPath());
    assertEquals(java.util.Optional.of(4), readRequest.jsonLine());
    assertEquals(java.util.Optional.of(12), readRequest.jsonColumn());

    dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest mergedRead =
        readRequest.withJson(
            dev.erst.gridgrind.contract.dto.ProblemContext.JsonLocation.located("ignored", 9, 22));
    assertEquals(java.util.Optional.of("steps[0]"), mergedRead.jsonPath());
    assertEquals(java.util.Optional.of(4), mergedRead.jsonLine());
    assertEquals(java.util.Optional.of(12), mergedRead.jsonColumn());

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep mergedExecute =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known("NEW", "NONE"),
                new dev.erst.gridgrind.contract.dto.ProblemContext.StepReference(
                    1, "cells", "INSPECTION", "GET_CELLS"),
                dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.unknown())
            .withLocation(
                dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.formulaCell(
                    "Budget", "A1", "SUM(A1)"));
    assertEquals(java.util.Optional.of("Budget"), mergedExecute.sheetName());
    assertEquals(java.util.Optional.of("A1"), mergedExecute.address());
    assertEquals(java.util.Optional.empty(), mergedExecute.range());
    assertEquals(java.util.Optional.of("SUM(A1)"), mergedExecute.formula());
    assertEquals(java.util.Optional.empty(), mergedExecute.namedRangeName());
  }

  @Test
  void cellReportDefaultAccessorsDispatchAcrossEverySubtype() {
    HyperlinkTarget hyperlink = new HyperlinkTarget.Url("https://example.com/budget");
    GridGrindResponse.CommentReport comment =
        new GridGrindResponse.CommentReport("Reviewed", "Alice", true);
    dev.erst.gridgrind.contract.dto.CellReport textCell =
        new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
            "A1",
            "STRING",
            "Reviewed",
            minimalStyle(),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            "Reviewed",
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.CellReport numberCell =
        new dev.erst.gridgrind.contract.dto.CellReport.NumberReport(
            "A2",
            "NUMERIC",
            "42",
            minimalStyle(),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            42.0d);
    dev.erst.gridgrind.contract.dto.CellReport booleanCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BooleanReport(
            "A3",
            "BOOLEAN",
            "TRUE",
            minimalStyle(),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            true);
    dev.erst.gridgrind.contract.dto.CellReport errorCell =
        new dev.erst.gridgrind.contract.dto.CellReport.ErrorReport(
            "A4",
            "ERROR",
            "#REF!",
            minimalStyle(),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            "#REF!");
    dev.erst.gridgrind.contract.dto.CellReport formulaCell =
        new dev.erst.gridgrind.contract.dto.CellReport.FormulaReport(
            "A5",
            "FORMULA",
            "42",
            minimalStyle(),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            "SUM(A2:A4)",
            new dev.erst.gridgrind.contract.dto.CellReport.NumberReport(
                "A5",
                "NUMERIC",
                "42",
                minimalStyle(),
                java.util.Optional.empty(),
                java.util.Optional.empty(),
                42.0d));

    assertEquals(java.util.Optional.of(hyperlink), textCell.hyperlink());
    assertEquals(java.util.Optional.of(comment), textCell.comment());
    assertEquals(java.util.Optional.of(hyperlink), numberCell.hyperlink());
    assertEquals(java.util.Optional.of(comment), numberCell.comment());
    assertEquals(java.util.Optional.of(hyperlink), booleanCell.hyperlink());
    assertEquals(java.util.Optional.of(comment), booleanCell.comment());
    assertEquals(java.util.Optional.of(hyperlink), errorCell.hyperlink());
    assertEquals(java.util.Optional.of(comment), errorCell.comment());
    assertEquals(java.util.Optional.of(hyperlink), formulaCell.hyperlink());
    assertEquals(java.util.Optional.of(comment), formulaCell.comment());
  }

  @Test
  void autofilterEntryDefaultSortStateDispatchesAcrossTableOwnedEntries() {
    AutofilterSortStateReport sortState =
        new AutofilterSortStateReport("A1:B4", false, true, "none", List.of());
    AutofilterEntryReport entry =
        new AutofilterEntryReport.TableOwned("A1:B4", "BudgetTable", List.of(), sortState);

    assertEquals(sortState, entry.sortState());
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
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known(
                        "NEW", "NONE")),
                java.util.Optional.empty(),
                null)
            .causes());
  }

  @Test
  void successCopiesAssertionsAndProblemCanCarryAssertionFailure() {
    List<AssertionResult> assertions = new ArrayList<>();
    assertions.add(new AssertionResult("assert-total", "EXPECT_CELL_VALUE"));

    GridGrindResponse.Success success =
        GridGrindResponses.success(
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
            new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known("NEW", "NONE"),
                new dev.erst.gridgrind.contract.dto.ProblemContext.StepReference(
                    0, "assert-total", "ASSERTION", "EXPECT_CELL_VALUE"),
                dev.erst.gridgrind.contract.dto.ProblemContext.ProblemLocation.cell(
                    "Budget", "B4")),
            java.util.Optional.of(failure),
            List.of());
    assertEquals(failure, problem.assertionFailure().orElseThrow());
  }

  @Test
  void executionJournalValidatesPhasesAndFailureClassification() {
    ExecutionJournal.Phase successPhase =
        new ExecutionJournal.Phase(
            ExecutionJournal.Status.SUCCEEDED,
            java.util.Optional.of("2026-04-18T10:00:00Z"),
            java.util.Optional.of("2026-04-18T10:00:01Z"),
            1);
    ExecutionJournal journal =
        new ExecutionJournal(
            java.util.Optional.of("budget-audit"),
            ExecutionJournalLevel.VERBOSE,
            new ExecutionJournal.SourceSummary(
                java.util.Optional.of("NEW"), java.util.Optional.empty()),
            new ExecutionJournal.PersistenceSummary(
                java.util.Optional.of("SAVE_AS"), java.util.Optional.of("/tmp/report.xlsx")),
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
                    java.util.Optional.of(
                        new ExecutionJournal.FailureClassification(
                            GridGrindProblemCode.ASSERTION_FAILED,
                            GridGrindProblemCategory.REQUEST,
                            "EXECUTE_STEP",
                            "observed value mismatch")))),
            List.of(),
            new ExecutionJournal.Outcome(
                ExecutionJournal.Status.FAILED,
                1,
                0,
                22,
                java.util.Optional.of(0),
                java.util.Optional.of("assert-total"),
                java.util.Optional.of(GridGrindProblemCode.ASSERTION_FAILED)),
            List.of(
                new ExecutionJournal.Event(
                    "2026-04-18T10:00:00Z",
                    "STEP",
                    "started",
                    java.util.Optional.of(0),
                    java.util.Optional.of("assert-total"))));

    assertEquals("budget-audit", journal.planId().orElseThrow());
    assertEquals(ExecutionJournalLevel.VERBOSE, journal.level());
    assertEquals("Cell Budget!B4", journal.steps().getFirst().resolvedTargets().getFirst().label());
    assertEquals(
        "NOT_STARTED phases must omit timestamps and use durationMillis=0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExecutionJournal.Phase(
                        ExecutionJournal.Status.NOT_STARTED,
                        java.util.Optional.of("2026-04-18T10:00:00Z"),
                        java.util.Optional.empty(),
                        0))
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
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }
}
