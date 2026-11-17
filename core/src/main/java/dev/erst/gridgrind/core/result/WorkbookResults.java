package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Public result factories that keep {@link WorkbookResult} focused on contract shapes. */
public final class WorkbookResults {
  private WorkbookResults() {}

  /**
   * Creates a successful response with a synthetic success journal and a not-requested calculation
   * report.
   */
  public static WorkbookResult.Success success(
      ProtocolVersion protocolVersion,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return WorkbookResultSupport.success(
        protocolVersion, persistence, warnings, assertions, inspections);
  }

  /** Creates one successful response using the current protocol version and no persistence. */
  public static WorkbookResult.Success success(
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return success(
        ProtocolVersion.CURRENT,
        new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
        warnings,
        assertions,
        inspections);
  }

  /**
   * Creates a failure response with a synthetic failure journal and a not-requested calculation
   * report.
   */
  public static WorkbookResult.Failure failure(
      ProtocolVersion protocolVersion,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      GridGrindProblemDetail.Problem problem) {
    return failure(protocolVersion, Optional.empty(), persistence, problem);
  }

  /** Creates a failure result with the request's optional stable plan identifier. */
  public static WorkbookResult.Failure failure(
      ProtocolVersion protocolVersion,
      Optional<PlanId> planId,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      GridGrindProblemDetail.Problem problem) {
    return WorkbookResultSupport.failure(protocolVersion, planId, persistence, problem);
  }

  /**
   * Creates a failure response with a synthetic failure journal and a not-requested calculation
   * report.
   */
  public static WorkbookResult.Failure failure(
      ProtocolVersion protocolVersion, GridGrindProblemDetail.Problem problem) {
    return failure(
        protocolVersion, new WorkbookResultPersistence.PersistenceOutcome.NotSaved(), problem);
  }

  /** Creates one failed response using the current protocol version. */
  public static WorkbookResult.Failure failure(GridGrindProblemDetail.Problem problem) {
    return failure(ProtocolVersion.CURRENT, problem);
  }

  /** Converts a completed result to failure while preserving established execution facts. */
  public static WorkbookResult.Failure withSecondaryFailure(
      WorkbookResult prior, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(prior, "prior must not be null");
    return withSecondaryFailure(prior, problem, prior.journal());
  }

  /** Retains explicit completed phase telemetry while attaching a secondary failure. */
  public static WorkbookResult.Failure withSecondaryFailure(
      WorkbookResult prior, GridGrindProblemDetail.Problem problem, ExecutionJournal journal) {
    Objects.requireNonNull(prior, "prior must not be null");
    Objects.requireNonNull(problem, "problem must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    var primary =
        prior instanceof WorkbookResult.Failure failure
            ? failure.problem().withCause(problem.asCause())
            : problem;
    Optional<ExecutionJournal.FailureStep> priorFailedStep =
        prior.journal().outcome() instanceof ExecutionJournal.Outcome.Failed failed
            ? failed.failedStep()
            : Optional.empty();
    ExecutionJournal failedJournal = failedJournal(journal, primary.code(), priorFailedStep);
    return switch (prior) {
      case WorkbookResult.Success success ->
          new WorkbookResult.Failure(
              success.protocolVersion(),
              success.planId(),
              failedJournal,
              success.calculation(),
              success.persistence(),
              success.evidence(),
              success.warnings(),
              success.assertions(),
              success.inspections(),
              primary);
      case WorkbookResult.Failure failure ->
          new WorkbookResult.Failure(
              failure.protocolVersion(),
              failure.planId(),
              failedJournal,
              failure.calculation(),
              failure.persistence(),
              failure.evidence(),
              failure.warnings(),
              failure.assertions(),
              failure.inspections(),
              primary);
    };
  }

  private static ExecutionJournal failedJournal(
      ExecutionJournal journal,
      dev.erst.gridgrind.core.problem.GridGrindProblemCode code,
      Optional<ExecutionJournal.FailureStep> priorFailedStep) {
    ExecutionJournal.Outcome.Failed outcome =
        switch (journal.outcome()) {
          case ExecutionJournal.Outcome.Succeeded success ->
              new ExecutionJournal.Outcome.Failed(
                  success.plannedStepCount(),
                  success.completedStepCount(),
                  success.durationMillis(),
                  code,
                  priorFailedStep);
          case ExecutionJournal.Outcome.Failed failure ->
              new ExecutionJournal.Outcome.Failed(
                  failure.plannedStepCount(),
                  failure.completedStepCount(),
                  failure.durationMillis(),
                  code,
                  failure.failedStep().or(() -> priorFailedStep));
        };
    return new ExecutionJournal(
        journal.level(),
        journal.source(),
        journal.validation(),
        journal.inputResolution(),
        journal.open(),
        journal.calculation(),
        journal.persistencePhase(),
        journal.close(),
        journal.steps(),
        outcome);
  }

  /**
   * Describes persistence that was requested but whose final destination was never touched.
   *
   * <p>This keeps every failure response truthful about the caller's save intent without inventing
   * an execution path for an artifact that was never written.
   */
  public static WorkbookResultPersistence.PersistenceOutcome notAttemptedPersistenceOutcome(
      WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
              saveAs.path(), new WorkbookResultPersistence.PublicationOutcome.NotAttempted());
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
        Optional<String> sourcePath =
            switch (request.source()) {
              case WorkbookPlan.WorkbookSource.New _ -> Optional.empty();
              case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
                  Optional.of(existingFile.path());
            };
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            sourcePath, new WorkbookResultPersistence.PublicationOutcome.NotAttempted());
      }
    };
  }

  /** Projects one proven or uncertain publication fact into the request's persistence intent. */
  public static WorkbookResultPersistence.PersistenceOutcome publicationPersistenceOutcome(
      WorkbookPlan request, WorkbookResultPersistence.PublicationOutcome publication) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(publication, "publication must not be null");
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          throw new IllegalArgumentException("NONE persistence cannot carry a publication outcome");
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new WorkbookResultPersistence.PersistenceOutcome.SavedAs(saveAs.path(), publication);
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
        Optional<String> sourcePath =
            switch (request.source()) {
              case WorkbookPlan.WorkbookSource.New _ -> Optional.empty();
              case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
                  Optional.of(existingFile.path());
            };
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(sourcePath, publication);
      }
    };
  }
}
