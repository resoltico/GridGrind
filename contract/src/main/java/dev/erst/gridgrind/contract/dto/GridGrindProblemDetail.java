package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Deterministic problem payloads returned for unsuccessful executions. */
public interface GridGrindProblemDetail {
  /** Deterministic failure payload returned for unsuccessful executions. */
  record Problem(
      GridGrindProblemCode code,
      GridGrindProblemCategory category,
      GridGrindProblemRecovery recovery,
      String title,
      String message,
      String resolution,
      ProblemContext context,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<AssertionFailure> assertionFailure,
      List<ProblemCause> causes) {
    /** Validates one deterministic problem payload and normalizes optional nested fields. */
    public Problem {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(category, "category must not be null");
      Objects.requireNonNull(recovery, "recovery must not be null");
      Objects.requireNonNull(title, "title must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(resolution, "resolution must not be null");
      Objects.requireNonNull(context, "context must not be null");
      assertionFailure = Objects.requireNonNullElseGet(assertionFailure, Optional::empty);
      causes = GridGrindResponseSupport.copyProblemCauses(causes);
    }

    /**
     * Constructs a Problem from a code and message, deriving category, recovery, title, and
     * resolution from the code.
     */
    public static Problem of(GridGrindProblemCode code, String message, ProblemContext context) {
      return new Problem(
          code,
          code.category(),
          code.recovery(),
          code.title(),
          message,
          code.resolution(),
          context,
          Optional.empty(),
          List.of());
    }
  }

  /** Individual GridGrind-classified diagnostic entries preserved for troubleshooting. */
  record ProblemCause(GridGrindProblemCode code, String message, String stage) {
    public ProblemCause {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(message, "message must not be null");
      Objects.requireNonNull(stage, "stage must not be null");
      if (message.isBlank()) {
        throw new IllegalArgumentException("message must not be blank");
      }
      if (stage.isBlank()) {
        throw new IllegalArgumentException("stage must not be blank");
      }
    }
  }
}
