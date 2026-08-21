package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;

/**
 * Orders phase-one request findings deterministically while keeping allocation ordinals internal.
 */
final class RequestStructuralProblemOrder {
  private RequestStructuralProblemOrder() {}

  static List<RequestStructuralProblem> order(List<RequestStructuralProblem> problems) {
    Objects.requireNonNull(problems, "problems must not be null");
    List<AllocatedProblem> allocated = new ArrayList<>(problems.size());
    for (int index = 0; index < problems.size(); index++) {
      allocated.add(new AllocatedProblem(index, problems.get(index)));
    }
    allocated.sort(
        Comparator.comparingInt(RequestStructuralProblemOrder::positionRank)
            .thenComparingLong(RequestStructuralProblemOrder::byteOffset)
            .thenComparingInt(RequestStructuralProblemOrder::stepIndex)
            .thenComparingInt(RequestStructuralProblemOrder::occurrenceOrdinal)
            .thenComparing(RequestStructuralProblemOrder::problemCode)
            .thenComparingInt(AllocatedProblem::ordinal));
    return allocated.stream().map(AllocatedProblem::problem).toList();
  }

  private static int positionRank(AllocatedProblem allocated) {
    return allocated.problem().byteOffset().isPresent() ? 0 : 1;
  }

  private static long byteOffset(AllocatedProblem allocated) {
    return allocated.problem().byteOffset().orElse(Long.MAX_VALUE);
  }

  private static int stepIndex(AllocatedProblem allocated) {
    return allocated.problem().jsonPath().map(RequestStructuralProblemOrder::stepIndex).orElse(-1);
  }

  private static int stepIndex(String jsonPath) {
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

  private static int occurrenceOrdinal(AllocatedProblem allocated) {
    return allocated.problem() instanceof RequestDuplicateKey duplicate
        ? duplicate.occurrenceOrdinal()
        : 0;
  }

  private static String problemCode(AllocatedProblem allocated) {
    return switch (Objects.requireNonNull(allocated.problem(), "problem must not be null")) {
      case RequestInvalidEncoding _ -> GridGrindProblemCode.INVALID_ENCODING.name();
      case RequestInvalidJson _ -> GridGrindProblemCode.INVALID_JSON.name();
      case RequestDuplicateKey _ -> GridGrindProblemCode.INVALID_JSON.name();
      case RequestNumberNotRepresentable _ -> GridGrindProblemCode.NUMBER_NOT_REPRESENTABLE.name();
      case RequestShapeStructuralProblem _ -> GridGrindProblemCode.INVALID_REQUEST_SHAPE.name();
    };
  }

  private record AllocatedProblem(int ordinal, RequestStructuralProblem problem) {
    private AllocatedProblem {
      // Allocation is private to order(), whose zero-based loop cannot create a negative ordinal.
      Objects.requireNonNull(problem, "problem must not be null");
    }
  }
}
