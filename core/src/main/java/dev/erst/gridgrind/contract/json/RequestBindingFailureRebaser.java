package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Reattaches one typed binding failure to a more precise authored request path. */
final class RequestBindingFailureRebaser {
  private RequestBindingFailureRebaser() {}

  static IllegalArgumentException rebase(
      IllegalArgumentException exception, String jsonPath, RuntimeException cause) {
    return switch (exception) {
      case FormulaRequestException formula ->
          new FormulaRequestException(
              formula.problemCode(),
              formula.getMessage(),
              Optional.of(jsonPath),
              Optional.empty(),
              Optional.empty(),
              cause);
      case InvalidRequestException invalid ->
          new InvalidRequestException(
              (RequestProblemDescriptor.Invariant) invalid.requestProblem(),
              Optional.of(jsonPath),
              Optional.empty(),
              Optional.empty(),
              cause);
      default -> exception;
    };
  }
}
