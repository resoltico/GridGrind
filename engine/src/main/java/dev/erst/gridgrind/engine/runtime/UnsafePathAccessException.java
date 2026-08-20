package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;

/** Raised when no-follow path binding cannot establish or retain a safe filesystem topology. */
final class UnsafePathAccessException extends IOException {
  private static final long serialVersionUID = 1L;

  UnsafePathAccessException(String message) {
    super(message);
  }

  UnsafePathAccessException(String message, Throwable cause) {
    super(message, cause);
  }
}
