package dev.erst.gridgrind.excel;

import java.net.URI;
import java.util.regex.Pattern;

/** Package-private non-throwing hyperlink target checks reused by authoring and analysis paths. */
final class ExcelHyperlinkValidation {
  private static final Pattern ABSOLUTE_URI_PATTERN =
      Pattern.compile("^[A-Za-z][A-Za-z0-9+.-]*:\\S+$");

  private ExcelHyperlinkValidation() {}

  static boolean isValidUrlTarget(String target) {
    if (target == null || target.isBlank() || !ABSOLUTE_URI_PATTERN.matcher(target).matches()) {
      return false;
    }
    try {
      return URI.create(target).isAbsolute();
    } catch (IllegalArgumentException exception) {
      return false;
    }
  }

  static boolean isValidEmailTarget(String target) {
    return target != null && !stripMailtoPrefix(target).isBlank();
  }

  static boolean isValidFileTarget(String target) {
    return target != null && !target.isBlank();
  }

  static boolean isValidDocumentTarget(String target) {
    return target != null && !target.isBlank();
  }

  static String stripMailtoPrefix(String email) {
    if (email.regionMatches(true, 0, "mailto:", 0, 7)) {
      return email.substring(7);
    }
    return email;
  }
}
