package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Protocol-owned scalar hyperlink target validation helpers. */
final class ProtocolHyperlinkTargetSupport {
  private ProtocolHyperlinkTargetSupport() {}

  static String normalizeDocumentTarget(String target) {
    return requireNonBlank(target, "target");
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static String stripMailtoPrefix(String email) {
    if (email.regionMatches(true, 0, "mailto:", 0, 7)) {
      return email.substring(7);
    }
    return email;
  }
}
