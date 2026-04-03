package dev.erst.gridgrind.protocol.dto;

import java.util.Locale;
import java.util.Objects;

/** Shared RGB color normalization helpers for protocol-facing style payloads. */
final class ProtocolRgbColorSupport {
  private ProtocolRgbColorSupport() {}

  /** Normalizes one optional {@code #RRGGBB} literal for protocol storage and comparison. */
  static String normalizeRgbHex(String color, String fieldName) {
    if (color == null) {
      return null;
    }
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    if (color.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    if (!color.matches("^#[0-9A-Fa-f]{6}$")) {
      throw new IllegalArgumentException(fieldName + " must match #RRGGBB");
    }
    return color.toUpperCase(Locale.ROOT);
  }
}
