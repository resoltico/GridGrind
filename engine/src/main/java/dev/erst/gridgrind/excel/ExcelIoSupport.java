package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.UncheckedIOException;

/** Shared unchecked IO bridge for workbook-side helper flows. */
public final class ExcelIoSupport {
  private ExcelIoSupport() {}

  /** Executes one checked-IO supplier and rethrows failures as {@link UncheckedIOException}. */
  public static <T> T unchecked(String failureMessage, IoSupplier<T> supplier) {
    try {
      return supplier.get();
    } catch (IOException exception) {
      throw new UncheckedIOException(failureMessage, exception);
    }
  }

  /** Checked-IO supplier contract for internal workbook helper flows. */
  @FunctionalInterface
  public interface IoSupplier<T> {
    /** Produces one value and may fail with checked IO. */
    T get() throws IOException;
  }
}
