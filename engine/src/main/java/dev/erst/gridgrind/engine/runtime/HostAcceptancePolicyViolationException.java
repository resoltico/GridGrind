package dev.erst.gridgrind.engine.runtime;

/** Signals that an authored execution mode cannot satisfy a trusted host acceptance requirement. */
final class HostAcceptancePolicyViolationException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  HostAcceptancePolicyViolationException(String message) {
    super(message);
  }
}
