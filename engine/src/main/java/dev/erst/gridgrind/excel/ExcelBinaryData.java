package dev.erst.gridgrind.excel;

import java.util.Arrays;
import java.util.Objects;

/** Immutable binary payload used by drawing and embedded-object workflows. */
public final class ExcelBinaryData {
  private final byte[] bytes;

  /** Creates an immutable binary payload from one non-empty byte array. */
  public ExcelBinaryData(byte[] bytes) {
    this(bytes, false);
  }

  /**
   * Creates an immutable binary payload from workbook-read bytes, which may be empty in malformed
   * packages that GridGrind still reports factually.
   */
  public static ExcelBinaryData readback(byte[] bytes) {
    return new ExcelBinaryData(bytes, true);
  }

  private ExcelBinaryData(byte[] bytes, boolean allowEmpty) {
    Objects.requireNonNull(bytes, "bytes must not be null");
    if (!allowEmpty && bytes.length == 0) {
      throw new IllegalArgumentException("bytes must not be empty");
    }
    this.bytes = bytes.clone();
  }

  /** Returns a defensive copy of the payload bytes. */
  public byte[] bytes() {
    return bytes.clone();
  }

  /** Returns the payload size in bytes. */
  public int size() {
    return bytes.length;
  }

  @Override
  public boolean equals(Object other) {
    return other instanceof ExcelBinaryData that && Arrays.equals(bytes, that.bytes);
  }

  @Override
  public int hashCode() {
    return Arrays.hashCode(bytes);
  }

  @Override
  public String toString() {
    return "ExcelBinaryData[size=" + bytes.length + "]";
  }
}
