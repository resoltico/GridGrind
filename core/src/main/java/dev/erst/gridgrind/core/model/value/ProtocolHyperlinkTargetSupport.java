package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Protocol-owned scalar hyperlink target validation helpers. */
final class ProtocolHyperlinkTargetSupport {
  private ProtocolHyperlinkTargetSupport() {}

  static String normalizeDocumentTarget(String target) {
    return TextValidation.requireNonBlank(target, "target");
  }

  static String stripMailtoPrefix(String email) {
    if (email.regionMatches(true, 0, "mailto:", 0, 7)) {
      return email.substring(7);
    }
    return email;
  }
}
