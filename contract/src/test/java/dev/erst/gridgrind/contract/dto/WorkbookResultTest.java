package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for step-based successful and failed protocol responses. */
class WorkbookResultTest {
  @Test
  void successDefaultsPersistenceAndCopiesWarningsAndInspections() {
    List<RequestWarning> warnings = new ArrayList<>();
    warnings.add(
        new RequestWarning(
            dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
            0,
            "set-total",
            "SET_CELL",
            "Quote spaced sheet names in formulas."));
    List<InspectionResult> inspections = new ArrayList<>();
    inspections.add(
        new WorkbookInspectionResult.WorkbookSummaryResult(
            "summary",
            new WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), 0, false)));

    WorkbookResult.Success success = WorkbookResults.success(warnings, List.of(), inspections);
    WorkbookResult.Success successWithoutInspections =
        WorkbookResults.success(warnings, List.of(), List.of());

    warnings.clear();
    inspections.clear();

    assertEquals(GridGrindProtocolVersion.current(), success.protocolVersion());
    assertInstanceOf(
        WorkbookResultPersistence.PersistenceOutcome.NotSaved.class, success.persistence());
    assertEquals(java.util.Optional.empty(), success.planId());
    assertEquals(java.util.Optional.empty(), success.journal().source().type());
    assertEquals(ExecutionJournal.Status.SUCCEEDED, success.journal().outcome().status());
    assertEquals(1, success.warnings().size());
    assertEquals(1, success.inspections().size());
    assertEquals(List.of(), successWithoutInspections.inspections());
  }

  @Test
  void failureBackfillConstructorCreatesFailedSyntheticJournal() {
    WorkbookResult.Failure failure =
        WorkbookResults.failure(
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "boom",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .unknown())));

    assertEquals(java.util.Optional.empty(), failure.planId());
    assertEquals(java.util.Optional.empty(), failure.journal().source().type());
    assertInstanceOf(
        WorkbookResultPersistence.PersistenceOutcome.NotSaved.class, failure.persistence());
    ExecutionJournal.Outcome.Failed outcome =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failure.journal().outcome());
    assertEquals(ExecutionJournal.Status.FAILED, outcome.status());
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, outcome.problemCode());
  }

  @Test
  void unwrittenPersistenceOutcomesPreserveRequestedSaveIntentWithoutInventingPaths() {
    WorkbookPlan overwriteExistingRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("fixtures/budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecurityInput.none()),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan saveAsRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "fixtures/output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan impossibleOverwriteRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecurityInput.none()),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of());

    WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            WorkbookResults.unwrittenPersistenceOutcome(overwriteExistingRequest));
    WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class,
            WorkbookResults.unwrittenPersistenceOutcome(saveAsRequest));
    WorkbookResultPersistence.PersistenceOutcome.Overwritten impossibleOverwrite =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            WorkbookResults.unwrittenPersistenceOutcome(impossibleOverwriteRequest));

    assertEquals(Optional.of("fixtures/budget.xlsx"), overwritten.sourcePath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, overwritten.write());
    assertEquals("fixtures/output.xlsx", savedAs.requestedPath());
    assertInstanceOf(WorkbookResultPersistence.WriteResult.NotWritten.class, savedAs.write());
    assertEquals(Optional.empty(), impossibleOverwrite.sourcePath());
    assertInstanceOf(
        WorkbookResultPersistence.WriteResult.NotWritten.class, impossibleOverwrite.write());
  }

  @Test
  void syntheticJournalRejectsInvalidFailureCodeCombinations() {
    IllegalArgumentException missingFailureCode =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookResult.syntheticJournal(
                    ExecutionJournal.Status.FAILED, java.util.Optional.empty()));
    assertEquals("FAILED outcomes must include failureCode", missingFailureCode.getMessage());

    IllegalArgumentException unexpectedFailureCode =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookResult.syntheticJournal(
                    ExecutionJournal.Status.SUCCEEDED,
                    java.util.Optional.of(GridGrindProblemCode.INVALID_ARGUMENTS)));
    assertEquals(
        "failureCode is only permitted when status is FAILED", unexpectedFailureCode.getMessage());

    IllegalArgumentException unsupportedNotStarted =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookResult.syntheticJournal(
                    ExecutionJournal.Status.NOT_STARTED, Optional.empty()));
    assertEquals(
        "synthetic journal outcome does not support NOT_STARTED",
        unsupportedNotStarted.getMessage());

    IllegalArgumentException unsupportedNotRequested =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookResult.syntheticJournal(
                    ExecutionJournal.Status.NOT_REQUESTED, Optional.empty()));
    assertEquals(
        "synthetic journal outcome does not support NOT_REQUESTED",
        unsupportedNotRequested.getMessage());
  }

  @Test
  void requestWarningsRequireStepIdentity() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                -1,
                "a",
                "SET_CELL",
                "warn"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                0,
                " ",
                "SET_CELL",
                "warn"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                0,
                "a",
                " ",
                "warn"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                0,
                "a",
                "SET_CELL",
                " "));
  }

  @Test
  void executeStepContextMergesTypedLocationsWithoutReintroducingNullPadding() {
    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep base =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                "EXISTING", "SAVE_AS"),
            new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference(
                2, "formula-health", "INSPECTION", "ANALYZE_FORMULA_HEALTH"),
            dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation.sheet(
                "Summary"));

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep enriched =
        base.withLocation(
            dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
                .formulaCell("Ignored", "B4", "SUM(B2:B3)"));

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
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                "--request"));
    GridGrindProblemDetail.Problem problem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown field 'reads'",
            new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE"),
                new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference(
                    1, "cells", "INSPECTION", "GET_CELLS"),
                dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation.cell(
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
    NamedRangeReport.FormulaReport formulaOnly =
        new NamedRangeReport.FormulaReport(
            "BudgetExpr", new NamedRangeScope.Workbook(), "SUM(Budget!A1:A3)");
    SheetProtectionReport.Unprotected unprotected = new SheetProtectionReport.Unprotected();
    dev.erst.gridgrind.contract.dto.CellReport.BlankReport blankCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BlankReport(
            "A1",
            java.util.Optional.of(""),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest readRequest =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput.requestFile(
                "/tmp/request.json"),
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation.located(
                "steps[0]", 4, 12));

    assertEquals("BudgetExpr", formulaOnly.name());
    assertEquals("SUM(Budget!A1:A3)", formulaOnly.refersToFormula());
    assertInstanceOf(SheetProtectionReport.Unprotected.class, unprotected);
    assertEquals("BLANK", blankCell.type());
    assertEquals(java.util.Optional.empty(), blankCell.hyperlink());
    assertEquals(java.util.Optional.empty(), blankCell.comment());
    assertEquals(Optional.empty(), new AutofilterEntryReport.SheetOwned("A1:B2").sortState());
    assertEquals(java.util.Optional.of("/tmp/request.json"), readRequest.requestPath());
    assertEquals(java.util.Optional.of("steps[0]"), readRequest.jsonPath());
    assertEquals(java.util.Optional.of(4), readRequest.jsonLine());
    assertEquals(java.util.Optional.of(12), readRequest.jsonColumn());

    dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest mergedRead =
        readRequest.withJson(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation.located(
                "ignored", 9, 22));
    assertEquals(java.util.Optional.of("steps[0]"), mergedRead.jsonPath());
    assertEquals(java.util.Optional.of(4), mergedRead.jsonLine());
    assertEquals(java.util.Optional.of(12), mergedRead.jsonColumn());

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep mergedExecute =
        new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE"),
                new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference(
                    1, "cells", "INSPECTION", "GET_CELLS"),
                dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
                    .unknown())
            .withLocation(
                dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
                    .formulaCell("Budget", "A1", "SUM(A1)"));
    assertEquals(java.util.Optional.of("Budget"), mergedExecute.sheetName());
    assertEquals(java.util.Optional.of("A1"), mergedExecute.address());
    assertEquals(java.util.Optional.empty(), mergedExecute.range());
    assertEquals(java.util.Optional.of("SUM(A1)"), mergedExecute.formula());
    assertEquals(java.util.Optional.empty(), mergedExecute.namedRangeName());
  }

  @Test
  void cellReportDefaultAccessorsDispatchAcrossEverySubtype() {
    HyperlinkTarget hyperlink = new HyperlinkTarget.Url("https://example.com/budget");
    CommentReport comment = new CommentReport("Reviewed", "Alice", true);
    dev.erst.gridgrind.contract.dto.CellReport textCell =
        new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
            "A1",
            java.util.Optional.of("Reviewed"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of("Reviewed"),
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.CellReport numberCell =
        new dev.erst.gridgrind.contract.dto.CellReport.NumberReport(
            "A2",
            java.util.Optional.of("42"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of(42.0d),
            java.util.Optional.empty());
    dev.erst.gridgrind.contract.dto.CellReport booleanCell =
        new dev.erst.gridgrind.contract.dto.CellReport.BooleanReport(
            "A3",
            java.util.Optional.of("TRUE"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of(true));
    dev.erst.gridgrind.contract.dto.CellReport errorCell =
        new dev.erst.gridgrind.contract.dto.CellReport.ErrorReport(
            "A4",
            java.util.Optional.of("#REF!"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of("#REF!"));
    dev.erst.gridgrind.contract.dto.CellReport formulaCell =
        new dev.erst.gridgrind.contract.dto.CellReport.FormulaReport(
            "A5",
            java.util.Optional.of("42"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of("SUM(A2:A4)"),
            java.util.Optional.of(
                new CellValueReport.NumberValue(42.0d, java.util.Optional.empty())));

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
        AutofilterSortStateReport.withoutSortMethod("A1:B4", false, true, List.of());
    AutofilterEntryReport entry =
        new AutofilterEntryReport.TableOwned(
            "A1:B4", "BudgetTable", List.of(), Optional.of(sortState));

    assertEquals(Optional.of(sortState), entry.sortState());
  }

  @Test
  void validatesPersistenceOutcomesWorkbookSummariesAndProblemCauses() {
    assertEquals(
        "requestedPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
                        " ", new WorkbookResultPersistence.WriteResult.Written("/tmp/out.xlsx")))
            .getMessage());
    assertEquals(
        "executionPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
                        "budget.xlsx", new WorkbookResultPersistence.WriteResult.Written(" ")))
            .getMessage());
    assertEquals(
        "sheetCount must be 0 for an empty workbook",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookSummary.Empty(1, List.of("Budget"), 0, false))
            .getMessage());
    assertEquals(
        "selectedSheetNames must only contain values present in sheetNames",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Ops"), 0, false))
            .getMessage());
    assertEquals(
        "message must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProblemDetail.ProblemCause(
                        GridGrindProblemCode.INVALID_REQUEST, " ", "READ_REQUEST"))
            .getMessage());
    assertEquals(
        "stage must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new GridGrindProblemDetail.ProblemCause(
                        GridGrindProblemCode.INVALID_REQUEST, "bad request", " "))
            .getMessage());
    assertEquals(
        List.of(),
        new GridGrindProblemDetail.Problem(
                GridGrindProblemCode.INVALID_REQUEST,
                GridGrindProblemCode.INVALID_REQUEST.category(),
                GridGrindProblemCode.INVALID_REQUEST.recovery(),
                GridGrindProblemCode.INVALID_REQUEST.title(),
                "bad request",
                GridGrindProblemCode.INVALID_REQUEST.resolution(),
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE")),
                List.of())
            .causes());
  }

  @Test
  void successCopiesAssertionsAndProblemCanCarryAssertionFailure() {
    List<AssertionResult> assertions = new ArrayList<>();
    assertions.add(new AssertionResult.Passed("assert-total", "EXPECT_CELL_VALUE"));

    WorkbookResult.Success success =
        WorkbookResults.success(
            List.of(),
            assertions,
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    "summary",
                    new WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 0, false))));
    assertions.clear();

    assertEquals(1, success.assertions().size());
    assertEquals("assert-total", success.assertions().getFirst().stepId());

    GridGrindProblemDetail.Problem problem =
        new GridGrindProblemDetail.Problem(
            GridGrindProblemCode.ASSERTION_FAILED,
            GridGrindProblemCode.ASSERTION_FAILED.category(),
            GridGrindProblemCode.ASSERTION_FAILED.recovery(),
            GridGrindProblemCode.ASSERTION_FAILED.title(),
            "assertion failed",
            GridGrindProblemCode.ASSERTION_FAILED.resolution(),
            new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE"),
                new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference(
                    0, "assert-total", "ASSERTION", "EXPECT_CELL_VALUE"),
                dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation.cell(
                    "Budget", "B4")),
            List.of());
    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, problem.code());
  }

  @Test
  void executionJournalValidatesPhasesAndFailureClassification() {
    ExecutionJournal.Phase successPhase =
        ExecutionJournal.Phase.succeeded("2026-04-18T10:00:00Z", "2026-04-18T10:00:01Z", 1);
    ExecutionJournal journal =
        new ExecutionJournal(
            ExecutionJournalLevel.VERBOSE,
            new ExecutionJournal.SourceSummary(
                java.util.Optional.of("NEW"), java.util.Optional.empty()),
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
            ExecutionJournal.Outcome.failed(
                1,
                0,
                22,
                GridGrindProblemCode.ASSERTION_FAILED,
                java.util.Optional.of(new ExecutionJournal.FailureStep(0, "assert-total"))));

    assertEquals(ExecutionJournalLevel.VERBOSE, journal.level());
    assertEquals("Cell Budget!B4", journal.steps().getFirst().resolvedTargets().getFirst().label());
    assertEquals(
        "2026-04-18T10:00:00Z",
        assertInstanceOf(ExecutionJournal.Phase.Succeeded.class, successPhase)
            .timing()
            .orElseThrow()
            .startedAt());
  }

  private static CellStyleReport minimalStyle() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
    return new CellStyleReport(
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
