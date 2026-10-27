package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliExecutionBindingsFactory;
import dev.erst.gridgrind.cli.render.PathGuidance;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/** Builds machine-readable doctor reports from raw request payloads before execution begins. */
final class CliDoctorRequestAnalyzer {
  private final GridGrindRequestDoctor requestDoctor;

  CliDoctorRequestAnalyzer(GridGrindRequestDoctor requestDoctor) {
    this.requestDoctor = GridGrindRequestDoctor.requireNonNull(requestDoctor);
  }

  RequestDoctorReport diagnose(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      RequestIntake analysis,
      InputStream stdin)
      throws IOException {
    return diagnose(
        requestPath,
        executionRootPath,
        tempRootPath,
        Optional.empty(),
        Optional.empty(),
        analysis,
        stdin);
  }

  RequestDoctorReport diagnose(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      Optional<Path> grantPath,
      Optional<Path> secretsProviderPath,
      RequestIntake analysis,
      InputStream stdin)
      throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(grantPath, "grantPath must not be null");
    Objects.requireNonNull(secretsProviderPath, "secretsProviderPath must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");

    ProblemContextRequestSurfaces.RequestInput requestInput =
        CliRequestReadFailureSupport.requestInput(requestPath);
    RequestDoctorReport baseReport =
        runBaseDoctorReport(
            requestPath,
            executionRootPath,
            tempRootPath,
            grantPath,
            secretsProviderPath,
            stdin,
            analysis,
            requestInput);
    if (!analysis.isBindable()) {
      return baseReport;
    }
    WorkbookPlan boundRequest = analysis.requireCompletePlan();
    if (CliPathArguments.isStandardInputPath(requestPath)
        && GridGrindRequestRequirements.requiresStandardInput(boundRequest)) {
      GridGrindProblemDetail.Problem standardInputProblem =
          dev.erst.gridgrind.engine.api.GridGrindProblems.problem(
              dev.erst.gridgrind.core.problem.GridGrindProblemCode.INVALID_REQUEST,
              PathGuidance.standardInputRequiresRequestMessage(),
              new dev.erst.gridgrind.core.problem.ProblemContext.ValidateRequest(
                  requestShape(boundRequest)),
              List.of());
      return RequestDoctorReport.invalid(
          baseReport.summary(),
          baseReport.warnings(),
          Stream.concat(Stream.of(standardInputProblem), baseReport.problems().stream()).toList());
    }
    return baseReport;
  }

  private RequestDoctorReport runBaseDoctorReport(
      Optional<Path> requestPath,
      Optional<Path> executionRootPath,
      Optional<Path> tempRootPath,
      Optional<Path> grantPath,
      Optional<Path> secretsProviderPath,
      InputStream stdin,
      RequestIntake analysis,
      ProblemContextRequestSurfaces.RequestInput requestInput)
      throws IOException {
    if (!analysis.isBindable()) {
      return requestDoctor.diagnose(analysis, requestInput);
    }
    if (grantPath.isEmpty()) {
      return requestDoctor.diagnose(analysis, requestInput);
    }
    WorkbookPlan request = analysis.requireCompletePlan();
    if (requestPath.or(() -> executionRootPath).isEmpty()) {
      return requestDoctor.diagnose(analysis, requestInput);
    }
    try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            requestPath,
            executionRootPath,
            tempRootPath,
            grantPath,
            secretsProviderPath,
            request,
            stdin)) {
      return requestDoctor.diagnose(analysis, requestInput, bindings.inputs());
    }
  }

  private static ProblemContextRequestSurfaces.RequestShape requestShape(WorkbookPlan request) {
    return ProblemContextRequestSurfaces.RequestShape.known(
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
