package dev.erst.gridgrind.engine.runtime;

/** Signals that a plan effect exceeds the authority supplied by the trusted execution host. */
final class ExecutionAuthorityDeniedException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  ExecutionAuthorityDeniedException(String message) {
    super(message);
  }
}
