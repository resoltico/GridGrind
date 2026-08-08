package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
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
      RequestAnalysis analysis,
      InputStream stdin)
      throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(executionRootPath, "executionRootPath must not be null");
    Objects.requireNonNull(tempRootPath, "tempRootPath must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");

    ProblemContextRequestSurfaces.RequestInput requestInput = requestInput(requestPath);
    List<GridGrindProblemDetail.Problem> intakeProblems =
        CliRequestAnalysisProblems.problems(analysis, requestInput);
    if (!intakeProblems.isEmpty()) {
      return RequestDoctorReport.invalid(Optional.empty(), List.of(), intakeProblems);
    }

    WorkbookPlan boundRequest = analysis.requireCompletePlan();
    RequestDoctorReport baseReport =
        runBaseDoctorReport(requestPath, executionRootPath, tempRootPath, stdin, boundRequest);
    if (requestPath.isEmpty() && GridGrindRequestRequirements.requiresStandardInput(boundRequest)) {
      GridGrindProblemDetail.Problem standardInputProblem =
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              GridGrindRequestSurfaceContractText.standardInputRequiresRequestMessage(),
              new ProblemContext.ValidateRequest(requestShape(boundRequest)),
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
      InputStream stdin,
      WorkbookPlan request)
      throws IOException {
    if (requestPath.isPresent() || executionRootPath.isPresent()) {
      try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
          CliExecutionBindingsFactory.create(
              requestPath, executionRootPath, tempRootPath, request, stdin)) {
        return requestDoctor.diagnose(request, bindings.inputs());
      }
    }
    return requestDoctor.diagnose(request);
  }

  private static ProblemContextRequestSurfaces.RequestInput requestInput(
      Optional<Path> requestPath) {
    return requestPath.isEmpty()
        ? ProblemContextRequestSurfaces.RequestInput.standardInput()
        : ProblemContextRequestSurfaces.RequestInput.requestFile(
            requestPath.orElseThrow().toAbsolutePath().toString());
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
