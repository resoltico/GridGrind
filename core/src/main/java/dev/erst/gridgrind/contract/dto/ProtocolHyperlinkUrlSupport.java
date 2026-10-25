package dev.erst.gridgrind.contract.dto;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Locale;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Pattern;

/** Protocol-owned URL and email hyperlink normalization helpers. */
final class ProtocolHyperlinkUrlSupport {
  private static final Pattern ABSOLUTE_URI_PATTERN =
      Pattern.compile("^[A-Za-z][A-Za-z0-9+.-]*:\\S+$");
  private static final Set<String> ALLOWED_URL_SCHEMES =
      Set.of("http", "https", "ftp", "ftps"); // LIM-028

  private ProtocolHyperlinkUrlSupport() {}

  static String normalizeUrlTarget(String target) {
    String normalized = ProtocolHyperlinkTargetSupport.requireNonBlank(target, "target");
    if (isValidUrlTarget(normalized)) {
      return normalized;
    }
    String scheme = absoluteUriScheme(normalized).orElse(null);
    if ("file".equalsIgnoreCase(scheme)) {
      throw new IllegalArgumentException("target uses file: scheme; use FILE hyperlinks instead");
    }
    if ("mailto".equalsIgnoreCase(scheme)) {
      throw new IllegalArgumentException(
          "target uses mailto: scheme; use EMAIL hyperlinks instead");
    }
    if (scheme != null) {
      throw new IllegalArgumentException(
          "target uses unsupported scheme '"
              + scheme
              + "'; only http, https, ftp, and ftps are allowed");
    }
    throw new IllegalArgumentException("target must be an absolute URI with a scheme");
  }

  static String normalizeEmailTarget(String email) {
    String normalized =
        ProtocolHyperlinkTargetSupport.stripMailtoPrefix(
            ProtocolHyperlinkTargetSupport.requireNonBlank(email, "target"));
    if (!isValidEmailTarget(normalized)) {
      throw new IllegalArgumentException("target must be an email address");
    }
    return normalized;
  }

  static boolean looksLikeAbsoluteUri(String path) {
    if (path.length() >= 3
        && Character.isLetter(path.charAt(0))
        && path.charAt(1) == ':'
        && (path.charAt(2) == '/' || path.charAt(2) == '\\')) {
      return false;
    }
    try {
      return new URI(path).isAbsolute();
    } catch (URISyntaxException exception) {
      return false;
    }
  }

  private static boolean isValidUrlTarget(String target) {
    if (!ABSOLUTE_URI_PATTERN.matcher(target).matches()) {
      return false;
    }
    try {
      URI uri = URI.create(target);
      return ALLOWED_URL_SCHEMES.contains(uri.getScheme().toLowerCase(Locale.ROOT));
    } catch (IllegalArgumentException exception) {
      return false;
    }
  }

  private static boolean isValidEmailTarget(String target) {
    if (target.isBlank() || target.contains(" ")) {
      return false;
    }
    int atIndex = target.indexOf('@');
    return atIndex > 0 && atIndex == target.lastIndexOf('@') && atIndex < target.length() - 1;
  }

  private static Optional<String> absoluteUriScheme(String target) {
    try {
      URI uri = URI.create(target);
      return uri.isAbsolute() ? Optional.ofNullable(uri.getScheme()) : Optional.empty();
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }
}
