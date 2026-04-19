package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.AssertionFailure;
import java.util.Objects;

/** Internal checked exception used to surface structured assertion mismatches through problems. */
final class AssertionFailedException extends Exception {
  private static final long serialVersionUID = 1L;

  private final transient AssertionFailure assertionFailure;

  AssertionFailedException(String message, AssertionFailure assertionFailure) {
    super(Objects.requireNonNull(message, "message must not be null"));
    this.assertionFailure =
        Objects.requireNonNull(assertionFailure, "assertionFailure must not be null");
  }

  AssertionFailure assertionFailure() {
    return assertionFailure;
  }
}
