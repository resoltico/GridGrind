package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.value.ExcelDateSystem;
import dev.erst.gridgrind.core.semantics.ExecutionConventions;
import java.util.List;
import java.util.Objects;

/** Explicit owner of one private provider realm's settings and caches. */
public final class NativeInvocationContext implements AutoCloseable {
  private final List<Runnable> teardown;
  private boolean closed;

  NativeInvocationContext(List<Runnable> teardown) {
    this.teardown = List.copyOf(teardown);
  }

  /** Enters only a provider realm owned by the caller, never an embedding host's POI copy. */
  public static NativeInvocationContext enterOwned(ExecutionConventions conventions) {
    Objects.requireNonNull(conventions);
    var context =
        new NativeInvocationContext(
            List.of(
                org.apache.poi.util.ThreadLocalUtil::clearAllThreadLocals,
                org.apache.xmlbeans.ThreadLocalUtil::clearAllThreadLocals,
                org.apache.poi.util.LocaleUtil::resetUserLocale,
                org.apache.poi.util.LocaleUtil::resetUserTimeZone));
    try {
      org.apache.poi.util.LocaleUtil.setUserLocale(
          java.util.Locale.forLanguageTag(conventions.localeTag()));
      org.apache.poi.util.LocaleUtil.setUserTimeZone(
          java.util.TimeZone.getTimeZone(conventions.civilDateBridgeOffset()));
      org.apache.poi.ss.usermodel.DateSystemContext.configure(
          ExcelDateSystem.WINDOWS_1900.maximumSerial(), ExcelDateSystem.MAC_1904.maximumSerial());
      return context;
    } catch (RuntimeException | Error failure) {
      context.closePreserving(failure);
      throw failure;
    }
  }

  void requireOpen() {
    if (closed) throw new IllegalStateException("Native invocation context is already closed");
  }

  @Override
  public void close() {
    if (closed) return;
    closed = true;
    @org.jspecify.annotations.Nullable Throwable first = null;
    for (Runnable action : teardown)
      try {
        action.run();
      } catch (Error failure) {
        if (first == null) first = failure;
        else if (first instanceof Error) suppress(first, failure);
        else {
          suppress(failure, first);
          first = failure;
        }
      } catch (RuntimeException failure) {
        if (first == null) first = failure;
        else suppress(first, failure);
      }
    if (first instanceof Error error) throw error;
    if (first instanceof RuntimeException runtime) throw runtime;
  }

  void closePreserving(Throwable primary) {
    try {
      close();
    } catch (RuntimeException | Error cleanup) {
      suppress(primary, cleanup);
    }
  }

  @SuppressWarnings({"ReferenceEquality", "PMD.CompareObjectsWithEquals"})
  // Throwable self-suppression is forbidden by identity.
  static void suppress(Throwable primary, Throwable secondary) {
    if (primary != secondary && primary.getSuppressed().length < 8)
      primary.addSuppressed(secondary);
  }
}
