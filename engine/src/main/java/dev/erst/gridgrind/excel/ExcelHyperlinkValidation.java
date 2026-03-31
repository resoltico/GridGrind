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
      URI uri = URI.create(target);
      String scheme = uri.getScheme();
      return !"file".equalsIgnoreCase(scheme) && !"mailto".equalsIgnoreCase(scheme);
    } catch (IllegalArgumentException exception) {
      return false;
    }
  }

  static boolean isValidEmailTarget(String target) {
    if (target == null) {
      return false;
    }
    String normalized = stripMailtoPrefix(target);
    if (normalized.isBlank() || normalized.contains(" ")) {
      return false;
    }
    int atIndex = normalized.indexOf('@');
    return atIndex > 0
        && atIndex == normalized.lastIndexOf('@')
        && atIndex < normalized.length() - 1;
  }

  static boolean isValidFileTarget(String target) {
    try {
      normalizeFileTarget(target);
      return true;
    } catch (IllegalArgumentException | NullPointerException exception) {
      return false;
    }
  }

  static boolean isValidDocumentTarget(String target) {
    return target != null && !target.isBlank();
  }

  static String normalizeFileTarget(String target) {
    return ExcelFileHyperlinkTargets.normalizePath(target);
  }

  static String absoluteUriScheme(String target) {
    if (target == null || target.isBlank()) {
      return null;
    }
    try {
      URI uri = URI.create(target);
      return uri.isAbsolute() ? uri.getScheme() : null;
    } catch (IllegalArgumentException exception) {
      return null;
    }
  }

  static String stripMailtoPrefix(String email) {
    if (email.regionMatches(true, 0, "mailto:", 0, 7)) {
      return email.substring(7);
    }
    return email;
  }
}
