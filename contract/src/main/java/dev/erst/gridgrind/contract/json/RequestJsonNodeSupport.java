package dev.erst.gridgrind.contract.json;

import java.util.Objects;

/** Shared invariants for tolerant-request syntax values. */
final class RequestJsonNodeSupport {
  private RequestJsonNodeSupport() {}

  static long requireByteOffset(long byteOffset) {
    if (byteOffset < 0) {
      throw new IllegalArgumentException("byteOffset must not be negative");
    }
    return byteOffset;
  }

  static <T> T requireValue(T value, String fieldName) {
    return Objects.requireNonNull(value, fieldName + " must not be null");
  }
}
