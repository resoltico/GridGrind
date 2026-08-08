package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.function.ToIntFunction;

/** Provides the protocol's deterministic ordering for emitted problems and warnings. */
public final class DiagnosticOrder {
  private DiagnosticOrder() {}

  /** Returns problems in the complete protocol order without exposing allocation metadata. */
  public static List<GridGrindProblemDetail.Problem> problems(
      List<GridGrindProblemDetail.Problem> problems) {
    Objects.requireNonNull(problems, "problems must not be null");
    List<Allocated<GridGrindProblemDetail.Problem>> allocated =
        allocate(problems, DiagnosticOrder::phase);
    allocated.sort(problemComparator());
    return allocated.stream().map(Allocated::value).toList();
  }

  /** Returns warnings in the complete protocol order without exposing allocation metadata. */
  public static List<RequestWarning> warnings(List<RequestWarning> warnings) {
    Objects.requireNonNull(warnings, "warnings must not be null");
    List<Allocated<RequestWarning>> allocated = allocate(warnings, _ -> STATIC_VALIDATION_PHASE);
    allocated.sort(warningComparator());
    return allocated.stream().map(Allocated::value).toList();
  }

  private static Comparator<Allocated<GridGrindProblemDetail.Problem>> problemComparator() {
    return Comparator.comparingInt(
            (Allocated<GridGrindProblemDetail.Problem> problem) -> phase(problem.value()))
        .thenComparingInt(problem -> positionRank(problem.value()))
        .thenComparingLong(problem -> byteOffset(problem.value()))
        .thenComparingInt(problem -> stepIndex(problem.value()))
        .thenComparingInt(problem -> occurrenceOrdinal(problem.value()))
        .thenComparing(problem -> problem.value().code().name())
        .thenComparingInt(Allocated::ordinal);
  }

  private static Comparator<Allocated<RequestWarning>> warningComparator() {
    return Comparator.comparingInt((Allocated<RequestWarning> _) -> STATIC_VALIDATION_PHASE)
        .thenComparingInt(_ -> UNPOSITIONED_RANK)
        .thenComparingLong(_ -> UNPOSITIONED_BYTE_OFFSET)
        .thenComparingInt(warning -> warning.value().stepIndex())
        .thenComparingInt(_ -> NON_DUPLICATE_OCCURRENCE)
        .thenComparing(warning -> warning.value().code().name())
        .thenComparingInt(Allocated::ordinal);
  }

  /**
   * Gives each phase its own deterministic final tiebreaker. The ordinal is deliberately internal:
   * it orders otherwise identical diagnostics without becoming a protocol field.
   */
  private static <T> List<Allocated<T>> allocate(List<T> values, ToIntFunction<T> phaseSelector) {
    Objects.requireNonNull(phaseSelector, "phaseSelector must not be null");
    List<Allocated<T>> allocated = new ArrayList<>(values.size());
    Map<Integer, Integer> nextOrdinalByPhase = new HashMap<>();
    for (T value : values) {
      T nonNullValue = Objects.requireNonNull(value, "values must not contain nulls");
      int phase = phaseSelector.applyAsInt(nonNullValue);
      int ordinal = nextOrdinalByPhase.getOrDefault(phase, 0);
      nextOrdinalByPhase.put(phase, Math.incrementExact(ordinal));
      allocated.add(new Allocated<>(ordinal, nonNullValue));
    }
    return allocated;
  }

  private static final int STATIC_VALIDATION_PHASE = 3;
  private static final int UNPOSITIONED_RANK = 1;
  private static final long UNPOSITIONED_BYTE_OFFSET = Long.MAX_VALUE;
  private static final int NON_DUPLICATE_OCCURRENCE = 0;

  private static int phase(GridGrindProblemDetail.Problem problem) {
    return switch (problem.context()) {
      case ProblemContext.ParseArguments _, ProblemContext.ReadRequest _ -> 1;
      case ProblemContext.ValidateRequest _ -> 3;
      case ProblemContext.ResolveInputs _,
          ProblemContext.OpenWorkbook _,
          ProblemContext.ExecuteCalculation.Preflight _ ->
          4;
      case ProblemContext.ExecuteCalculation.Execution _,
          ProblemContext.ExecuteStep _,
          ProblemContext.PersistWorkbook _,
          ProblemContext.ExecuteRequest _,
          ProblemContext.WriteResponse _ ->
          5;
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
