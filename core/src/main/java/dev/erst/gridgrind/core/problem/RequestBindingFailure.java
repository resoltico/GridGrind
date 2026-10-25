package dev.erst.gridgrind.core.problem;

import java.util.Objects;
import java.util.Optional;

/**
 * One constructor-level request failure found while independently binding a valid JSON fragment.
 */
public record RequestBindingFailure(
    IllegalArgumentException exception, String jsonPath, Optional<Long> byteOffset) {
  public RequestBindingFailure {
    Objects.requireNonNull(exception, "exception must not be null");
    jsonPath = requireNonBlank(jsonPath, "jsonPath");
    byteOffset = Objects.requireNonNullElseGet(byteOffset, Optional::empty);
    byteOffset.ifPresent(
        offset -> {
          if (offset < 0) {
            throw new IllegalArgumentException("byteOffset must not be negative");
          }
        });
  }

  /** Returns this failure with its exact request-token location when the raw tree retains one. */
  public RequestBindingFailure locatedAt(Optional<Long> byteOffset) {
    return new RequestBindingFailure(exception, jsonPath, byteOffset);
  }

  /** Returns a copy whose typed problem and location identify a qualified request path. */
  public RequestBindingFailure rebasedAt(String qualifiedPath, Optional<Long> byteOffset) {
    IllegalArgumentException rebased =
        switch (exception) {
          case InvalidRequestException invalid ->
              new InvalidRequestException(
                  (RequestProblemDescriptor.Invariant) invalid.requestProblem(),
                  Optional.of(qualifiedPath),
                  Optional.empty(),
                  Optional.empty(),
                  invalid);
          case FormulaRequestException formula ->
              new FormulaRequestException(
                  formula.problemCode(),
                  formula.getMessage(),
                  Optional.of(qualifiedPath),
                  Optional.empty(),
                  Optional.empty(),
                  formula);
          default -> exception;
        };
    return new RequestBindingFailure(rebased, qualifiedPath, byteOffset);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
