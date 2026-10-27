package dev.erst.gridgrind.core.source;

import java.util.Arrays;
import java.util.Objects;

/** Immutable binary payload used by drawing and embedded-object workflows. */
public final class BinaryContent {
  private final byte[] bytes;

  /** Creates an immutable binary payload from one non-empty byte array. */
  public BinaryContent(byte[] bytes) {
    this(bytes, false);
  }

  /**
   * Creates an immutable binary payload from workbook-read bytes, which may be empty in malformed
   * packages that GridGrind still reports factually.
   */
  public static BinaryContent readback(byte[] bytes) {
    return new BinaryContent(bytes, true);
  }

  private BinaryContent(byte[] bytes, boolean allowEmpty) {
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
    return other instanceof BinaryContent that && Arrays.equals(bytes, that.bytes);
  }

  @Override
  public int hashCode() {
    return Arrays.hashCode(bytes);
  }

  @Override
  public String toString() {
    return "BinaryContent[size=" + bytes.length + "]";
  }
}
