package dev.erst.gridgrind.core.result;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.value.CellValueReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCategory;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedNameLabel;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
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
            dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
            0,
            new StepId("set-total"),
            "SET_CELL",
            "Quote spaced sheet names in formulas."));
    List<InspectionResult> inspections = new ArrayList<>();
    inspections.add(
        new WorkbookInspectionResult.WorkbookSummaryResult(
            new StepId("summary"),
            new WorkbookSummary.WithSheets(
                1,
                List.of(new SheetLabel("Budget")),
                new SheetLabel("Budget"),
                List.of(new SheetLabel("Budget")),
                0,
                false)));

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
                new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                    dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
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
  void notAttemptedPersistenceOutcomesPreserveRequestedSaveIntentWithoutInventingPaths() {
    WorkbookPlan overwriteExistingRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("fixtures/budget.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecuritySpec.none()),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan saveAsRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "fixtures/output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan impossibleOverwriteRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecuritySpec.none()),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());

    WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            WorkbookResults.notAttemptedPersistenceOutcome(overwriteExistingRequest));
    WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class,
            WorkbookResults.notAttemptedPersistenceOutcome(saveAsRequest));
    WorkbookResultPersistence.PersistenceOutcome.Overwritten impossibleOverwrite =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            WorkbookResults.notAttemptedPersistenceOutcome(impossibleOverwriteRequest));

    assertEquals(Optional.of("fixtures/budget.xlsx"), overwritten.sourcePath());
    assertInstanceOf(
        WorkbookResultPersistence.PublicationOutcome.NotAttempted.class, overwritten.publication());
    assertEquals("fixtures/output.xlsx", savedAs.requestedPath());
    assertInstanceOf(
        WorkbookResultPersistence.PublicationOutcome.NotAttempted.class, savedAs.publication());
    assertEquals(Optional.empty(), impossibleOverwrite.sourcePath());
    assertInstanceOf(
        WorkbookResultPersistence.PublicationOutcome.NotAttempted.class,
        impossibleOverwrite.publication());
  }

  @Test
  void publicationPersistenceOutcomesRetainTheRequestedDestinationAndRejectMemoryOnlyPlans() {
    WorkbookPlan noneRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan saveAsRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "fixtures/output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan overwriteRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("fixtures/input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecuritySpec.none()),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookPlan overwriteNewRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecuritySpec.none()),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            List.of());
    WorkbookResultPersistence.PublicationOutcome publication =
        new WorkbookResultPersistence.PublicationOutcome.NotPublished(
            WorkbookResultPersistence.PublicationOutcome.DestinationState.PRESERVED);

    WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.SavedAs.class,
            WorkbookResults.publicationPersistenceOutcome(saveAsRequest, publication));
    WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            WorkbookResults.publicationPersistenceOutcome(overwriteRequest, publication));
    WorkbookResultPersistence.PersistenceOutcome.Overwritten overwrittenNew =
        assertInstanceOf(
            WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
            WorkbookResults.publicationPersistenceOutcome(overwriteNewRequest, publication));

    assertEquals("fixtures/output.xlsx", savedAs.requestedPath());
    assertEquals(Optional.of("fixtures/input.xlsx"), overwritten.sourcePath());
    assertEquals(Optional.empty(), overwrittenNew.sourcePath());
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookResults.publicationPersistenceOutcome(noneRequest, publication));
  }

  @Test
  void publicationFactsRejectInvalidDigestAndNegativeByteCounts() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookResultPersistence.PublicationOutcome.Published(
                "output.xlsx",
                "invalid",
                0,
                new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
                new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
                    .FileAndDirectorySynced()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookResultPersistence.PublicationOutcome.Published(
                "output.xlsx",
                "0".repeat(64),
                -1,
                new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
                new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
                    .FileAndDirectorySynced()));
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
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                -1,
                new StepId("a"),
                "SET_CELL",
                "warn"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                0,
                new StepId(" "),
                "SET_CELL",
                "warn"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                0,
                new StepId("a"),
                " ",
                "warn"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new RequestWarning(
                dev.erst.gridgrind.core.problem.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                0,
                new StepId("a"),
                "SET_CELL",
                " "));
  }

  @Test
  void executeStepContextMergesTypedLocationsWithoutReintroducingNullPadding() {
    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep base =
        new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.known(
                "EXISTING", "SAVE_AS"),
            new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference(
                2, new StepId("formula-health"), "INSPECTION", "ANALYZE_FORMULA_HEALTH"),
            dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation.sheet(
                new SheetLabel("Summary")));

    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep enriched =
        base.withLocation(
            dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation
                .formulaCell(new SheetLabel("Ignored"), "B4", "SUM(B2:B3)"));

    assertEquals(
        java.util.Optional.of("Summary"),
        enriched.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(java.util.Optional.of("B4"), enriched.address());
    assertEquals(java.util.Optional.empty(), enriched.range());
    assertEquals(java.util.Optional.of("SUM(B2:B3)"), enriched.formula());
    assertEquals(java.util.Optional.empty(), enriched.namedRangeName());
  }

  @Test
  void parseArgumentsAndProblemsExposeStepCentricContext() {
    dev.erst.gridgrind.core.problem.ProblemContext.ParseArguments parseArguments =
        new dev.erst.gridgrind.core.problem.ProblemContext.ParseArguments(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.CliArgument.named(
                "--request"));
    GridGrindProblemDetail.Problem problem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST_SHAPE,
            "Unknown field 'reads'",
            new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE"),
                new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference(
                    1, new StepId("cells"), "INSPECTION", "GET_CELLS"),
                dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation.cell(
                    new SheetLabel("Budget"), "A1")));
    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep executeStep =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class, problem.context());

    assertEquals("PARSE_ARGUMENTS", parseArguments.stage());
    assertEquals(java.util.Optional.of("--request"), parseArguments.argumentName());
    assertEquals("EXECUTE_STEP", executeStep.stage());
    assertEquals(1, executeStep.stepIndex());
    assertEquals("cells", executeStep.stepId().value());
    assertEquals("GET_CELLS", executeStep.stepType());
    assertEquals(
        java.util.Optional.of("Budget"),
        executeStep.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(java.util.Optional.of("A1"), executeStep.address());
    assertEquals(java.util.Optional.empty(), executeStep.range());
    assertTrue(problem.title().contains("request"));
  }

  @Test
  void typedVariantsReplaceNullPaddingAndContextMergersPreserveExistingValues() {
    NamedRangeReport.FormulaReport formulaOnly =
        new NamedRangeReport.FormulaReport(
            new DefinedNameLabel("BudgetExpr"),
            new NamedRangeScope.Workbook(),
            "SUM(Budget!A1:A3)");
    SheetProtectionReport.Unprotected unprotected = new SheetProtectionReport.Unprotected();
    dev.erst.gridgrind.core.model.value.CellReport.BlankReport blankCell =
        new dev.erst.gridgrind.core.model.value.CellReport.BlankReport(
            new CellAddress("A1"),
            java.util.Optional.of(""),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.empty(),
            java.util.Optional.empty());
    dev.erst.gridgrind.core.problem.ProblemContext.ReadRequest readRequest =
        new dev.erst.gridgrind.core.problem.ProblemContext.ReadRequest(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestInput.requestFile(
                "/tmp/request.json"),
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation.located(
                "steps[0]", 4, 12));

    assertEquals("BudgetExpr", formulaOnly.name().value());
    assertEquals("SUM(Budget!A1:A3)", formulaOnly.refersToFormula());
    assertInstanceOf(SheetProtectionReport.Unprotected.class, unprotected);
    assertEquals("BLANK", blankCell.type());
    assertEquals(java.util.Optional.empty(), blankCell.hyperlink());
    assertEquals(java.util.Optional.empty(), blankCell.comment());
    assertEquals(
        Optional.empty(),
        new AutofilterEntryReport.SheetOwned(new RangeLabel("A1:B2")).sortState());
    assertEquals(java.util.Optional.of("/tmp/request.json"), readRequest.requestPath());
    assertEquals(java.util.Optional.of("steps[0]"), readRequest.jsonPath());
    assertEquals(java.util.Optional.of(4), readRequest.jsonLine());
    assertEquals(java.util.Optional.of(12), readRequest.jsonColumn());

    dev.erst.gridgrind.core.problem.ProblemContext.ReadRequest mergedRead =
        readRequest.withJson(
            dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation.located(
                "ignored", 9, 22));
    assertEquals(java.util.Optional.of("steps[0]"), mergedRead.jsonPath());
    assertEquals(java.util.Optional.of(4), mergedRead.jsonLine());
    assertEquals(java.util.Optional.of(12), mergedRead.jsonColumn());

    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep mergedExecute =
        new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE"),
                new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference(
                    1, new StepId("cells"), "INSPECTION", "GET_CELLS"),
                dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation
                    .unknown())
            .withLocation(
                dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation
                    .formulaCell(new SheetLabel("Budget"), "A1", "SUM(A1)"));
    assertEquals(
        java.util.Optional.of("Budget"),
        mergedExecute.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(java.util.Optional.of("A1"), mergedExecute.address());
    assertEquals(java.util.Optional.empty(), mergedExecute.range());
    assertEquals(java.util.Optional.of("SUM(A1)"), mergedExecute.formula());
    assertEquals(java.util.Optional.empty(), mergedExecute.namedRangeName());
  }

  @Test
  void cellReportDefaultAccessorsDispatchAcrossEverySubtype() {
    HyperlinkTarget hyperlink = new HyperlinkTarget.Url("https://example.com/budget");
    CommentReport comment = new CommentReport("Reviewed", "Alice", true);
    dev.erst.gridgrind.core.model.value.CellReport textCell =
        new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
            new CellAddress("A1"),
            java.util.Optional.of("Reviewed"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of("Reviewed"),
            java.util.Optional.empty());
    dev.erst.gridgrind.core.model.value.CellReport numberCell =
        new dev.erst.gridgrind.core.model.value.CellReport.NumberReport(
            new CellAddress("A2"),
            java.util.Optional.of("42"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of(42.0d),
            java.util.Optional.empty());
    dev.erst.gridgrind.core.model.value.CellReport booleanCell =
        new dev.erst.gridgrind.core.model.value.CellReport.BooleanReport(
            new CellAddress("A3"),
            java.util.Optional.of("TRUE"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of(true));
    dev.erst.gridgrind.core.model.value.CellReport errorCell =
        new dev.erst.gridgrind.core.model.value.CellReport.ErrorReport(
            new CellAddress("A4"),
            java.util.Optional.of("#REF!"),
            java.util.Optional.of(minimalStyle()),
            java.util.Optional.of(hyperlink),
            java.util.Optional.of(comment),
            java.util.Optional.of("#REF!"));
    dev.erst.gridgrind.core.model.value.CellReport formulaCell =
        new dev.erst.gridgrind.core.model.value.CellReport.FormulaReport(
            new CellAddress("A5"),
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
        AutofilterSortStateReport.withoutSortMethod(
            new RangeLabel("A1:B4"), false, true, List.of());
    AutofilterEntryReport entry =
        new AutofilterEntryReport.TableOwned(
            new RangeLabel("A1:B4"), "BudgetTable", List.of(), Optional.of(sortState));

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
                        " ", published("/tmp/out.xlsx")))
            .getMessage());
    assertEquals(
        "executionPath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
                        "budget.xlsx", published(" ")))
            .getMessage());
    assertEquals(
        "sheetCount must be 0 for an empty workbook",
        assertThrows(
                IllegalArgumentException.class,
                () -> new WorkbookSummary.Empty(1, List.of(new SheetLabel("Budget")), 0, false))
            .getMessage());
    assertEquals(
        "selectedSheetNames must only contain values present in sheetNames",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel("Budget"),
                        List.of(new SheetLabel("Ops")),
                        0,
                        false))
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
                new dev.erst.gridgrind.core.problem.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE")),
                List.of())
            .causes());
  }

  private static WorkbookResultPersistence.PublicationOutcome.Published published(
      String executionPath) {
    return new WorkbookResultPersistence.PublicationOutcome.Published(
        executionPath,
        "0".repeat(64),
        0,
        new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
        new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
            .FileSyncedDirectoryUnestablished());
  }

  @Test
  void successCopiesAssertionsAndProblemCanCarryAssertionFailure() {
    List<AssertionResult> assertions = new ArrayList<>();
    assertions.add(new AssertionResult.Passed(new StepId("assert-total"), "EXPECT_CELL_VALUE"));

    WorkbookResult.Success success =
        WorkbookResults.success(
            List.of(),
            assertions,
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    new StepId("summary"),
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel("Budget"),
                        List.of(new SheetLabel("Budget")),
                        0,
                        false))));
    assertions.clear();

    assertEquals(1, success.assertions().size());
    assertEquals("assert-total", success.assertions().getFirst().stepId().value());

    GridGrindProblemDetail.Problem problem =
        new GridGrindProblemDetail.Problem(
            GridGrindProblemCode.ASSERTION_FAILED,
            GridGrindProblemCode.ASSERTION_FAILED.category(),
            GridGrindProblemCode.ASSERTION_FAILED.recovery(),
            GridGrindProblemCode.ASSERTION_FAILED.title(),
            "assertion failed",
            GridGrindProblemCode.ASSERTION_FAILED.resolution(),
            new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE"),
                new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.StepReference(
                    0, new StepId("assert-total"), "ASSERTION", "EXPECT_CELL_VALUE"),
                dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation.cell(
                    new SheetLabel("Budget"), "B4")),
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
                    new StepId("assert-total"),
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
                java.util.Optional.of(
                    new ExecutionJournal.FailureStep(0, new StepId("assert-total")))));

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
