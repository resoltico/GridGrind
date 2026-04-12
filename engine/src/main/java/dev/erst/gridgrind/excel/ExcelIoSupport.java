package dev.erst.gridgrind.excel;

import java.io.IOException;

/** Shared unchecked IO bridge for workbook-side helper flows. */
final class ExcelIoSupport {
  private ExcelIoSupport() {}

  /** Executes one checked-IO supplier and rethrows failures as {@link IllegalStateException}. */
  static <T> T unchecked(String failureMessage, IoSupplier<T> supplier) {
    try {
      return supplier.get();
    } catch (IOException exception) {
      throw new IllegalStateException(failureMessage, exception);
    }
  }

  /** Checked-IO supplier contract for internal workbook helper flows. */
  @FunctionalInterface
  interface IoSupplier<T> {
    /** Produces one value and may fail with checked IO. */
    T get() throws IOException;
  }
}
