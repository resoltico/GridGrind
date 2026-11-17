package dev.erst.gridgrind.core.result;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Secondary failures change the final outcome without erasing established workbook facts. */
class SecondaryFailureRetentionTest {
  @Test
  void successfulPublicationRetainsEveryResultFactWhenHostCleanupFails() {
    var prior = completed();
    var cleanup = problem(GridGrindProblemCode.INTERNAL_ERROR, "synthetic cleanup failure");
    var failed = WorkbookResults.withSecondaryFailure(prior, cleanup);
    assertSame(cleanup, failed.problem());
    assertFacts(prior, failed);
    assertEquals(
        new ExecutionJournal.Outcome.Failed(2, 2, 47, cleanup.code(), Optional.empty()),
        failed.journal().outcome());
    assertEquals(prior.journal().steps(), failed.journal().steps());
    assertSame(prior.journal().close(), failed.journal().close());
    assertSame(prior.journal().persistencePhase(), failed.journal().persistencePhase());
  }

  @Test
  void anExistingPrimaryFailureAndFailedStepRemainPrimary() {
    var completed = completed();
    var primary = problem(GridGrindProblemCode.INVALID_REQUEST, "synthetic primary failure");
    var failedStep = new ExecutionJournal.FailureStep(2, new StepId("later"));
    var journal =
        journal(
            new ExecutionJournal.Outcome.Failed(3, 2, 51, primary.code(), Optional.of(failedStep)),
            ExecutionJournal.Phase.succeededWithoutTiming());
    var prior =
        new WorkbookResult.Failure(
            completed.protocolVersion(),
            completed.planId(),
            journal,
            completed.calculation(),
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            completed.evidence(),
            completed.warnings(),
            completed.assertions(),
            completed.inspections(),
            primary);
    var cleanup = problem(GridGrindProblemCode.INTERNAL_ERROR, "synthetic secondary failure");
    var failed = WorkbookResults.withSecondaryFailure(prior, cleanup);
    assertEquals(primary.code(), failed.problem().code());
    assertEquals(primary.message(), failed.problem().message());
    assertEquals(primary.context(), failed.problem().context());
    assertEquals(List.of(cleanup.asCause()), failed.problem().causes());
    assertEquals(journal.outcome(), failed.journal().outcome());
    assertEquals(prior.assertions(), failed.assertions());
    assertEquals(prior.inspections(), failed.inspections());
    assertSame(prior.evidence(), failed.evidence());
    var updatedClose =
        journal(
            new ExecutionJournal.Outcome.Failed(3, 2, 63, cleanup.code(), Optional.empty()),
            ExecutionJournal.Phase.failedWithoutTiming());
    var retained = WorkbookResults.withSecondaryFailure(prior, cleanup, updatedClose);
    assertEquals(
        new ExecutionJournal.Outcome.Failed(3, 2, 63, primary.code(), Optional.of(failedStep)),
        retained.journal().outcome());
  }

  @Test
  void explicitCloseTelemetryIsPreservedAndRepeatedSecondaryDiagnosticsAreBounded() {
    var prior = completed();
    var cleanup = problem(GridGrindProblemCode.INTERNAL_ERROR, "first cleanup failure");
    var close = ExecutionJournal.Phase.failedWithoutTiming();
    var completedClose = journal(new ExecutionJournal.Outcome.Succeeded(2, 2, 59), close);
    WorkbookResult.Failure failed =
        WorkbookResults.withSecondaryFailure(prior, cleanup, completedClose);
    for (int index = 0; index < 20; index++) {
      failed =
          WorkbookResults.withSecondaryFailure(
              failed, problem(GridGrindProblemCode.INTERNAL_ERROR, "additional cleanup " + index));
    }
    assertFacts(prior, failed);
    assertSame(close, failed.journal().close());
    assertEquals(
        new ExecutionJournal.Outcome.Failed(2, 2, 59, cleanup.code(), Optional.empty()),
        failed.journal().outcome());
    assertEquals(cleanup.message(), failed.problem().message());
    assertEquals(8, failed.problem().causes().size());
    assertTrue(failed.problem().causes().getFirst().message().contains("additional cleanup 0"));
  }

  private static WorkbookResult.Success completed() {
    var assertion = new AssertionResult.Passed(new StepId("verify"), "EXPECT_SHEET_PRESENT");
    var evidence =
        new WorkbookExecutionEvidence(
            WorkbookExecutionEvidence.Admission.empty(),
            new WorkbookExecutionEvidence.Structural.Established(),
            new WorkbookExecutionEvidence.Computational.NotRequested(),
            new WorkbookExecutionEvidence.TaskSpecific.Established(
                List.of(
                    new WorkbookExecutionEvidence.TaskSpecific.Assertion(
                        WorkbookExecutionEvidence.TaskSpecific.Origin.PLAN_AUTHORED, assertion))),
            new WorkbookExecutionEvidence.Presentational.NotAssessed(),
            List.of(new WorkbookExecutionEvidence.Preservation.NotRequested()));
    var publication =
        new WorkbookResultPersistence.PublicationOutcome.Published(
            "/synthetic/output.xlsx",
            "a".repeat(64),
            123,
            new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
            new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
                .FileSyncedDirectoryUnestablished());
    return new WorkbookResult.Success(
        ProtocolVersion.CURRENT,
        Optional.of(new PlanId("run")),
        journal(
            new ExecutionJournal.Outcome.Succeeded(2, 2, 47),
            ExecutionJournal.Phase.succeededWithoutTiming()),
        CalculationReport.notRequested(),
        new WorkbookResultPersistence.PersistenceOutcome.SavedAs("output.xlsx", publication),
        evidence,
        List.of(RequestWarning.progressSinkFailed()),
        List.of(assertion),
        List.of(
            new WorkbookInspectionResult.SheetsResult(
                new StepId("inspect"), List.of(new SheetLabel("Budget")))));
  }

  private static ExecutionJournal journal(
      ExecutionJournal.Outcome outcome, ExecutionJournal.Phase close) {
    var phase = ExecutionJournal.Phase.succeededWithoutTiming();
    return new ExecutionJournal(
        ExecutionJournalLevel.SUMMARY,
        new ExecutionJournal.SourceSummary(Optional.of("NEW"), Optional.empty()),
        phase,
        phase,
        phase,
        new ExecutionJournal.Calculation(
            ExecutionJournal.Phase.notRequested(), ExecutionJournal.Phase.notRequested()),
        phase,
        close,
        List.of(
            new ExecutionJournal.Step(
                0,
                new StepId("verify"),
                "ASSERTION",
                "EXPECT_SHEET_PRESENT",
                List.of(),
                phase,
                ExecutionJournal.StepOutcome.SUCCEEDED,
                Optional.empty()),
            new ExecutionJournal.Step(
                1,
                new StepId("inspect"),
                "INSPECTION",
                "GET_SHEETS",
                List.of(),
                phase,
                ExecutionJournal.StepOutcome.SUCCEEDED,
                Optional.empty())),
        outcome);
  }

  private static GridGrindProblemDetail.Problem problem(GridGrindProblemCode code, String message) {
    return GridGrindProblemDetail.Problem.of(
        code, message, new ProblemContext.ExecuteRequest(RequestShape.known("NEW", "SAVE_AS")));
  }

  private static void assertFacts(WorkbookResult.Success prior, WorkbookResult.Failure failed) {
    assertEquals(prior.protocolVersion(), failed.protocolVersion());
    assertEquals(prior.planId(), failed.planId());
    assertSame(prior.persistence(), failed.persistence());
    assertSame(prior.calculation(), failed.calculation());
    assertSame(prior.evidence(), failed.evidence());
    assertEquals(prior.warnings(), failed.warnings());
    assertEquals(prior.assertions(), failed.assertions());
    assertEquals(prior.inspections(), failed.inspections());
  }
}
