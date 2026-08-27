package dev.erst.gridgrind.contract.json;

import java.util.Objects;
import java.util.Optional;

/**
 * One constructor-level request failure found while independently binding a valid JSON fragment.
 */
public record RequestBindingFailure(
    IllegalArgumentException exception, String jsonPath, Optional<Long> byteOffset) {
  /** Normalizes one binding failure into the request-problem exception family. */
  static RequestBindingFailure from(
      RuntimeException failure,
      RequestJsonNode fragment,
      String fragmentPath,
      RequestDiagnosticRedactor redactor) {
    Objects.requireNonNull(failure, "failure must not be null");
    Objects.requireNonNull(fragment, "fragment must not be null");
    Objects.requireNonNull(fragmentPath, "fragmentPath must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    return RequestBindingFailureNormalizer.from(failure, fragment, fragmentPath, redactor);
  }

  /** Retains a complete-plan invariant at its already-qualified product-owned request path. */
  static RequestBindingFailure fromCompletePlan(
      RuntimeException failure, RequestDiagnosticRedactor redactor) {
    Objects.requireNonNull(failure, "failure must not be null");
    Objects.requireNonNull(redactor, "redactor must not be null");
    return RequestBindingFailureNormalizer.fromCompletePlan(failure, redactor);
  }

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
  RequestBindingFailure locatedAt(Optional<Long> byteOffset) {
    return new RequestBindingFailure(exception, jsonPath, byteOffset);
  }

  RequestBindingFailure rebasedAt(String qualifiedPath, Optional<Long> byteOffset) {
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
