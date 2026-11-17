package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliExecutionBindingsFactory;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResults;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Builds truthful failed execution results once the executor boundary has been crossed. */
final class CliExecutionFailureSupport {
  private CliExecutionFailureSupport() {}

  /**
   * Creates execution inputs, then retains the failed-result contract through binding cleanup.
   *
   * <p>Input construction occurs before the execution boundary and propagates its {@link
   * IOException}. Once the binding exists, ordinary exceptions, including a binding close failure,
   * are workbook execution failures rather than rejected commands. A missing executor result cannot
   * establish that publication was unattempted.
   */
  static WorkbookResult executeStarted(
      GridGrindRequestExecutor requestExecutor,
      WorkbookPlan request,
      RequestIntake analysis,
      ExecutionInputs executionInputs)
      throws IOException {
    Objects.requireNonNull(requestExecutor, "requestExecutor must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
    ExecutionInputs inputs =
        Objects.requireNonNull(executionInputs, "executionInputs must not be null");
    CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            inputs.requestPath(),
            inputs.executionRootPath(),
            inputs.tempRootPath(),
            inputs.grantPath(),
            inputs.secretsProviderPath(),
            request,
            inputs.stdin());
    return executeBound(
        request,
        bindings.inputs().workingDirectory(),
        bindings,
        () ->
            Objects.requireNonNull(
                requestExecutor.execute(analysis, bindings.inputs(), inputs.progressSink()),
                "requestExecutor must not return null"));
  }

  /** Owns entered execution and its binding until the actual result and cleanup are known. */
  static WorkbookResult executeBound(
      WorkbookPlan request,
      Path workingDirectory,
      AutoCloseable bindings,
      java.util.function.Supplier<WorkbookResult> execution) {
    WorkbookResult result = null;
    try (bindings) {
      result = Objects.requireNonNull(execution.get(), "executor result must not be null");
    } catch (Exception exception) {
      return result == null
          ? unreportedExecutionFailure(request, workingDirectory, exception)
          : WorkbookResults.withSecondaryFailure(
              result,
              GridGrindProblems.fromException(
                  exception, new ProblemContext.ExecuteRequest(requestShape(request))));
    }
    return result;
  }

  static WorkbookResult.Failure failureBeforeExecution(WorkbookPlan request, Throwable exception) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return WorkbookResults.failure(
        request.protocolVersion(),
        request.planId(),
        WorkbookResults.notAttemptedPersistenceOutcome(request),
        GridGrindProblems.fromException(
            exception, new ProblemContext.ExecuteRequest(requestShape(request))));
  }

  private static WorkbookResult.Failure unreportedExecutionFailure(
      WorkbookPlan request, Path workingDirectory, Throwable exception) {
    var persistence =
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ ->
              new dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome
                  .NotSaved();
          case WorkbookPlan.WorkbookPersistence.SaveAs save ->
              new dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome
                  .SavedAs(save.path(), uncertain(save.path(), workingDirectory));
          case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
            if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile source)) {
              throw new IllegalStateException(
                  "Executor failed without a resolvable publication target", exception);
            }
            yield new dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome
                .Overwritten(source.path(), uncertain(source.path(), workingDirectory));
          }
        };
    return WorkbookResults.failure(
        request.protocolVersion(),
        request.planId(),
        persistence,
        GridGrindProblems.fromException(
            exception, new ProblemContext.ExecuteRequest(requestShape(request))));
  }

  private static dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome
      uncertain(String rawPath, Path workingDirectory) {
    Path path = Path.of(rawPath);
    String resolved =
        (path.isAbsolute() ? path : workingDirectory.resolve(path))
            .toAbsolutePath()
            .normalize()
            .toString();
    return new dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome
        .Uncertain(
        resolved,
        dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome.Recovery
            .INSPECT_DESTINATION_DO_NOT_RETRY,
        dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome
            .CommitObservation.NOT_ESTABLISHED);
  }

  private static RequestShape requestShape(WorkbookPlan request) {
    return RequestShape.known(
        switch (request.source()) {
          case WorkbookPlan.WorkbookSource.New _ -> "NEW";
          case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
        },
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
          case WorkbookPlan.WorkbookPersistence.Overwrite _ -> "OVERWRITE";
          case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
        });
  }

  record ExecutionInputs(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      Optional<Path> grantPath,
      Optional<Path> secretsProviderPath,
      InputStream stdin,
      GridGrindProgressSink progressSink) {
    ExecutionInputs {
      Objects.requireNonNull(requestPath, "requestPath must not be null");
      Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
      Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
      Objects.requireNonNull(grantPath, "grantPath must not be null");
      Objects.requireNonNull(secretsProviderPath, "secretsProviderPath must not be null");
      Objects.requireNonNull(stdin, "stdin must not be null");
      Objects.requireNonNull(progressSink, "progressSink must not be null");
    }
  }
}
