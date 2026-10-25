package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Shared RGB color normalization helpers for protocol-facing style payloads. */
public final class ProtocolRgbColorSupport {
  private ProtocolRgbColorSupport() {}

  static ProtocolRgbColorSupport newForVerification() {
    return new ProtocolRgbColorSupport();
  }

  /** Normalizes one optional {@code #RRGGBB} literal for protocol storage and comparison. */
  public static Optional<String> normalizeRgbHex(String color, String fieldName) {
    if (color == null) {
      return Optional.empty();
    }
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    if (color.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    if (!color.matches(ProtocolConstraintValues.RGB_HEX_PATTERN)) {
      throw new IllegalArgumentException(fieldName + " must match #RRGGBB");
    }
    return Optional.of(color.toUpperCase(Locale.ROOT));
  }

  /** Normalizes one required {@code #RRGGBB} literal for protocol storage and comparison. */
  public static String requireRgbHex(String color, String fieldName) {
    return normalizeRgbHex(
            Objects.requireNonNull(color, fieldName + " must not be null"), fieldName)
        .orElseThrow();
  }
}
