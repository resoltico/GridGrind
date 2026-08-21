package dev.erst.gridgrind.engine.runtime;

/** Raised when a request-owned path resolves outside the configured execution root. */
final class RequestPathEscapeException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  RequestPathEscapeException(String message) {
    super(message);
  }
}
