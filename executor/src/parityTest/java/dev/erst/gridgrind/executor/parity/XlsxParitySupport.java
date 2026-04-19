package dev.erst.gridgrind.executor.parity;

import java.util.Objects;
import java.util.concurrent.Callable;

/** Shared exception-wrapping support for the parity ledger, corpus, and oracle harness. */
final class XlsxParitySupport {
  private XlsxParitySupport() {}

  static <T> T call(String action, Callable<T> callable) {
    Objects.requireNonNull(action, "action must not be null");
    Objects.requireNonNull(callable, "callable must not be null");
    try {
      return callable.call();
    } catch (RuntimeException exception) {
      throw exception;
    } catch (Exception exception) {
      throw new XlsxParityException(action, exception);
    }
  }
}

/** Runtime failure wrapper used when parity harness work cannot complete successfully. */
final class XlsxParityException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  XlsxParityException(String action, Exception cause) {
    super(action, cause);
  }
}
