package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestBindingFailure;
import dev.erst.gridgrind.contract.json.RequestDuplicateKey;
import dev.erst.gridgrind.contract.json.RequestStructuralProblem;
import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;

/** Projects all tolerant-intake findings into the canonical public problem core. */
public final class GridGrindRequestAnalysisProblems {
  private GridGrindRequestAnalysisProblems() {}

  /** Returns every structural and binding problem in their original request-analysis order. */
  public static List<GridGrindProblemDetail.Problem> project(
      RequestAnalysis analysis, ProblemContextRequestSurfaces.RequestInput requestInput) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    return Stream.concat(
            analysis.structuralProblems().stream()
                .map(problem -> structuralProblem(problem, requestInput)),
            analysis.bindingFailures().stream()
                .map(problem -> bindingProblem(problem, requestInput)))
        .toList();
  }

  private static GridGrindProblemDetail.Problem structuralProblem(
      RequestStructuralProblem problem, ProblemContextRequestSurfaces.RequestInput requestInput) {
    return GridGrindProblems.fromException(
        GridGrindJson.structuralException(problem),
        new ProblemContext.ReadRequest(requestInput, locationFor(problem)));
  }

  private static GridGrindProblemDetail.Problem bindingProblem(
      RequestBindingFailure problem, ProblemContextRequestSurfaces.RequestInput requestInput) {
    return GridGrindProblems.fromException(
        problem.exception(), new ProblemContext.BindRequest(requestInput, locationFor(problem)));
  }

  private static ProblemContextRequestSurfaces.JsonLocation locationFor(
      RequestBindingFailure problem) {
    return locationFor(problem.jsonPath(), problem.byteOffset());
  }

  static ProblemContextRequestSurfaces.JsonLocation locationFor(
      String jsonPath, java.util.Optional<Long> byteOffset) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    Objects.requireNonNull(byteOffset, "byteOffset must not be null");
    return byteOffset
        .map(
            offset -> ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(jsonPath, offset))
        .orElseGet(() -> ProblemContextRequestSurfaces.JsonLocation.pathOnly(jsonPath));
  }

  private static ProblemContextRequestSurfaces.JsonLocation locationFor(
      RequestStructuralProblem problem) {
    if (problem instanceof RequestDuplicateKey duplicate) {
      return ProblemContextRequestSurfaces.JsonLocation.duplicateKey(
          duplicate.containingObjectPath(),
          duplicate.key(),
          duplicate.occurrenceOrdinal(),
          duplicate.tokenByteOffset());
    }
    java.util.Optional<String> jsonPath = problem.jsonPath();
    java.util.Optional<Long> byteOffset = problem.byteOffset();
    if (jsonPath.isPresent() && byteOffset.isPresent()) {
      return ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
          jsonPath.orElseThrow(), byteOffset.orElseThrow());
    }
    if (jsonPath.isPresent()) {
      return ProblemContextRequestSurfaces.JsonLocation.pathOnly(jsonPath.orElseThrow());
    }
    return ProblemContextRequestSurfaces.JsonLocation.byteOffset(byteOffset.orElseThrow());
  }
}
