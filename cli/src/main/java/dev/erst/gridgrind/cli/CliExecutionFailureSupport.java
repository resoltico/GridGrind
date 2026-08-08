package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.engine.api.GridGrindJournalSink;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
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
   * IOException}. Once the binding exists, every failure, including a binding close failure, is a
   * workbook execution failure rather than a rejected command.
   */
  static WorkbookResult executeStarted(
      GridGrindRequestExecutor requestExecutor,
      WorkbookPlan request,
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      InputStream stdin,
      GridGrindJournalSink journalSink)
      throws IOException {
    Objects.requireNonNull(requestExecutor, "requestExecutor must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    Objects.requireNonNull(journalSink, "journalSink must not be null");
    CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            requestPath, executionRootPath, tempRootPath, request, stdin);
    try (bindings) {
      return Objects.requireNonNull(
          requestExecutor.execute(request, bindings.inputs(), journalSink),
          "requestExecutor must not return null");
    } catch (Throwable exception) {
      return failure(request, exception);
    }
  }

  static WorkbookResult.Failure failure(WorkbookPlan request, Throwable exception) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return WorkbookResults.failure(
        request.protocolVersion(),
        request.planId(),
        WorkbookResults.unwrittenPersistenceOutcome(request),
        GridGrindProblems.fromException(
            exception, new ProblemContext.ExecuteRequest(requestShape(request))));
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
}
