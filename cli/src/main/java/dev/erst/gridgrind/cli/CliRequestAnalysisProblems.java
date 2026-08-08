package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestBindingFailure;
import dev.erst.gridgrind.contract.json.RequestDuplicateKey;
import dev.erst.gridgrind.contract.json.RequestStructuralProblem;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Stream;

/** Converts all ordered request-intake findings into the one public problem projection. */
final class CliRequestAnalysisProblems {
  private CliRequestAnalysisProblems() {}

  static List<GridGrindProblemDetail.Problem> problems(
      RequestAnalysis analysis, ProblemContextRequestSurfaces.RequestInput requestInput) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    return Stream.concat(
            analysis.structuralProblems().stream()
                .map(problem -> LocatedProblem.structural(problem, requestInput)),
            analysis.bindingFailures().stream()
                .map(problem -> LocatedProblem.binding(problem, requestInput)))
        .sorted(Comparator.comparingLong(LocatedProblem::byteOffset))
        .map(LocatedProblem::problem)
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
        problem.exception(), new ProblemContext.ReadRequest(requestInput, locationFor(problem)));
  }

  static ProblemContextRequestSurfaces.JsonLocation locationFor(RequestBindingFailure problem) {
    return problem
        .byteOffset()
        .map(
            byteOffset ->
                ProblemContextRequestSurfaces.JsonLocation.pathAtByteOffset(
                    problem.jsonPath(), byteOffset))
        .orElseGet(() -> ProblemContextRequestSurfaces.JsonLocation.pathOnly(problem.jsonPath()));
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

  private record LocatedProblem(long byteOffset, GridGrindProblemDetail.Problem problem) {
    private LocatedProblem {
      Objects.requireNonNull(problem, "problem must not be null");
    }

    static LocatedProblem structural(
        RequestStructuralProblem problem, ProblemContextRequestSurfaces.RequestInput requestInput) {
      return new LocatedProblem(
          problem.byteOffset().orElse(Long.MAX_VALUE), structuralProblem(problem, requestInput));
    }

    static LocatedProblem binding(
        RequestBindingFailure problem, ProblemContextRequestSurfaces.RequestInput requestInput) {
      return new LocatedProblem(
          problem.byteOffset().orElse(Long.MAX_VALUE), bindingProblem(problem, requestInput));
    }
  }
}
