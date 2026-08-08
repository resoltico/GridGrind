package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Provides the protocol's deterministic ordering for emitted problems and warnings. */
public final class DiagnosticOrder {
  private DiagnosticOrder() {}

  /** Returns problems in the complete protocol order without exposing allocation metadata. */
  public static List<GridGrindProblemDetail.Problem> problems(
      List<GridGrindProblemDetail.Problem> problems) {
    Objects.requireNonNull(problems, "problems must not be null");
    List<Allocated<GridGrindProblemDetail.Problem>> allocated = allocate(problems);
    allocated.sort(problemComparator());
    return allocated.stream().map(Allocated::value).toList();
  }

  /** Returns warnings in the complete protocol order without exposing allocation metadata. */
  public static List<RequestWarning> warnings(List<RequestWarning> warnings) {
    Objects.requireNonNull(warnings, "warnings must not be null");
    List<Allocated<RequestWarning>> allocated = allocate(warnings);
    allocated.sort(
        Comparator.comparingInt((Allocated<RequestWarning> _) -> 3)
            .thenComparingInt(_ -> 1)
            .thenComparingLong(_ -> Long.MAX_VALUE)
            .thenComparingInt(warning -> warning.value().stepIndex())
            .thenComparingInt(_ -> 0)
            .thenComparing(warning -> warning.value().code().name())
            .thenComparingInt(Allocated::ordinal));
    return allocated.stream().map(Allocated::value).toList();
  }

  private static Comparator<Allocated<GridGrindProblemDetail.Problem>> problemComparator() {
    return Comparator.comparingInt((Allocated<GridGrindProblemDetail.Problem> problem) -> phase(problem.value()))
        .thenComparingInt(problem -> positionRank(problem.value()))
        .thenComparingLong(problem -> byteOffset(problem.value()))
        .thenComparingInt(problem -> stepIndex(problem.value()))
        .thenComparingInt(problem -> occurrenceOrdinal(problem.value()))
        .thenComparing(problem -> problem.value().code().name())
        .thenComparingInt(Allocated::ordinal);
  }

  private static <T> List<Allocated<T>> allocate(List<T> values) {
    List<Allocated<T>> allocated = new ArrayList<>(values.size());
    for (int ordinal = 0; ordinal < values.size(); ordinal++) {
      allocated.add(new Allocated<>(ordinal, Objects.requireNonNull(values.get(ordinal), "values must not contain nulls")));
    }
    return allocated;
  }

  private static int phase(GridGrindProblemDetail.Problem problem) {
    return switch (problem.context()) {
      case ProblemContext.ParseArguments _, ProblemContext.ReadRequest _ -> 1;
      case ProblemContext.ValidateRequest _ -> 3;
      case ProblemContext.ResolveInputs _,
          ProblemContext.OpenWorkbook _,
          ProblemContext.ExecuteCalculation.Preflight _ -> 4;
      case ProblemContext.ExecuteCalculation.Execution _,
          ProblemContext.ExecuteStep _,
          ProblemContext.PersistWorkbook _,
          ProblemContext.ExecuteRequest _,
          ProblemContext.WriteResponse _ -> 5;
    };
  }

  private static int positionRank(GridGrindProblemDetail.Problem problem) {
    return byteOffsetOptional(problem).isPresent() ? 0 : 1;
  }

  private static long byteOffset(GridGrindProblemDetail.Problem problem) {
    return byteOffsetOptional(problem).orElse(Long.MAX_VALUE);
  }

  private static Optional<Long> byteOffsetOptional(GridGrindProblemDetail.Problem problem) {
    return switch (problem.context()) {
      case ProblemContext.ReadRequest readRequest -> readRequest.byteOffset();
      default -> Optional.empty();
    };
  }

  private static int stepIndex(GridGrindProblemDetail.Problem problem) {
    return switch (problem.context()) {
      case ProblemContext.ExecuteStep executeStep -> executeStep.stepIndex();
      case ProblemContext.ReadRequest readRequest ->
          readRequest.jsonPath().map(DiagnosticOrder::stepIndexFromJsonPath).orElse(-1);
      default -> -1;
    };
  }

  private static int stepIndexFromJsonPath(String jsonPath) {
    if (!jsonPath.startsWith("steps[")) {
      return -1;
    }
    int closeBracket = jsonPath.indexOf(']');
    if (closeBracket < 0) {
      return -1;
    }
    try {
      return Integer.parseInt(jsonPath.substring("steps[".length(), closeBracket));
    } catch (NumberFormatException ignored) {
      return -1;
    }
  }

  private static int occurrenceOrdinal(GridGrindProblemDetail.Problem problem) {
    return switch (problem.context()) {
      case ProblemContext.ReadRequest readRequest ->
          readRequest.duplicateKey().map(JsonLocation.DuplicateKey::occurrenceOrdinal).orElse(0);
      default -> 0;
    };
  }

  /** Internal deterministic allocation carrier; never part of a serialized DTO. */
  private record Allocated<T>(int ordinal, T value) {
    private Allocated {
      if (ordinal < 0) {
        throw new IllegalArgumentException("ordinal must not be negative");
      }
      Objects.requireNonNull(value, "value must not be null");
    }
  }
}
