package dev.erst.gridgrind.contract.dto;

import java.net.URI;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.regex.Pattern;

/** Protocol-owned hyperlink normalization and validation helpers. */
final class ProtocolHyperlinkSupport {
  private static final Pattern ABSOLUTE_URI_PATTERN =
      Pattern.compile("^[A-Za-z][A-Za-z0-9+.-]*:\\S+$");
  private static final char[] HEX_DIGITS = "0123456789ABCDEF".toCharArray();
  private static final String SAFE_URI_PATH_CHARACTERS =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~/:";

  private ProtocolHyperlinkSupport() {}

  static String normalizeUrlTarget(String target) {
    String normalized = requireNonBlank(target, "target");
    if (isValidUrlTarget(normalized)) {
      return normalized;
    }
    String scheme = absoluteUriScheme(normalized);
    if ("file".equalsIgnoreCase(scheme)) {
      throw new IllegalArgumentException("target uses file: scheme; use FILE hyperlinks instead");
    }
    if ("mailto".equalsIgnoreCase(scheme)) {
      throw new IllegalArgumentException(
          "target uses mailto: scheme; use EMAIL hyperlinks instead");
    }
    throw new IllegalArgumentException("target must be an absolute URI with a scheme");
  }

  static String normalizeEmailTarget(String email) {
    String normalized = stripMailtoPrefix(requireNonBlank(email, "target"));
    if (!isValidEmailTarget(normalized)) {
      throw new IllegalArgumentException("target must be an email address");
    }
    return normalized;
  }

  static String normalizeFileTarget(String path) {
    Objects.requireNonNull(path, "path must not be null");
    if (path.isBlank()) {
      throw new IllegalArgumentException("path must not be blank");
    }
    if (looksLikeFileUri(path)) {
      return normalizeFileUri(path);
    }
    if (looksLikeAbsoluteUri(path)) {
      throw new IllegalArgumentException("path must be a local file path or file: URI");
    }
    return decodeEscapedRelativePath(path);
  }

  static String normalizeDocumentTarget(String target) {
    return requireNonBlank(target, "target");
  }

  static String toPoiFileAddress(String path) {
    String normalizedPath = normalizeFileTarget(path);
    if (looksLikeWindowsDrivePath(normalizedPath)) {
      return encodeAbsoluteUriPath("/" + normalizedPath.replace('\\', '/'));
    }

    Path candidate;
    try {
      candidate = Path.of(normalizedPath);
    } catch (InvalidPathException exception) {
      throw new IllegalArgumentException(invalidPathMessage(exception), exception);
    }

    if (candidate.isAbsolute()) {
      return candidate.toUri().toASCIIString();
    }
    return encodeRelativePath(normalizedPath);
  }

  private static boolean isValidUrlTarget(String target) {
    if (!ABSOLUTE_URI_PATTERN.matcher(target).matches()) {
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

  private static boolean isValidEmailTarget(String target) {
    if (target.isBlank() || target.contains(" ")) {
      return false;
    }
    int atIndex = target.indexOf('@');
    return atIndex > 0 && atIndex == target.lastIndexOf('@') && atIndex < target.length() - 1;
  }

  private static String absoluteUriScheme(String target) {
    try {
      URI uri = URI.create(target);
      return uri.isAbsolute() ? uri.getScheme() : null;
    } catch (IllegalArgumentException exception) {
      return null;
    }
  }

  private static boolean looksLikeFileUri(String path) {
    return path.regionMatches(true, 0, "file:", 0, 5);
  }

  private static String normalizeFileUri(String path) {
    try {
      URI uri = new URI(path);
      String authority = uri.getAuthority();
      String uriPath = Objects.requireNonNullElse(uri.getPath(), "");
      if (authority != null) {
        if (uriPath.isBlank()) {
          throw new IllegalArgumentException("path must contain a file-system path");
        }
        return "//" + authority + uriPath;
      }
      try {
        return Path.of(uri).toString();
      } catch (RuntimeException exception) {
        throw new IllegalArgumentException("path must be a valid file: URI", exception);
      }
    } catch (URISyntaxException exception) {
      throw new IllegalArgumentException("path must be a valid file: URI", exception);
    }
  }

  private static boolean looksLikeAbsoluteUri(String path) {
    if (looksLikeWindowsDrivePath(path)) {
      return false;
    }
    try {
      return new URI(path).isAbsolute();
    } catch (URISyntaxException exception) {
      return false;
    }
  }

  private static boolean looksLikeWindowsDrivePath(String path) {
    if (path.length() < 3) {
      return false;
    }
    char drive = path.charAt(0);
    char colon = path.charAt(1);
    char separator = path.charAt(2);
    if (!Character.isLetter(drive) || colon != ':') {
      return false;
    }
    return separator == '/' || separator == '\\';
  }

  private static String decodeEscapedRelativePath(String path) {
    if (!path.contains("%")) {
      return path;
    }
    try {
      URI uri = new URI(path);
      String decodedPath = Objects.requireNonNullElse(uri.getPath(), "");
      return decodedPath.isBlank() ? path : decodedPath;
    } catch (URISyntaxException exception) {
      return path;
    }
  }

  private static String encodeAbsoluteUriPath(String path) {
    return "file://" + encodePath(path);
  }

  private static String encodeRelativePath(String path) {
    return encodePath(path.replace('\\', '/'));
  }

  private static String encodePath(String path) {
    StringBuilder builder = new StringBuilder(path.length());
    path.codePoints()
        .forEach(
            codePoint -> {
              if (isPlainUriPathCodePoint(codePoint)) {
                builder.appendCodePoint(codePoint);
                return;
              }
              appendPercentEncoded(builder, codePoint);
            });
    return builder.toString();
  }

  private static boolean isPlainUriPathCodePoint(int codePoint) {
    return SAFE_URI_PATH_CHARACTERS.indexOf(codePoint) >= 0;
  }

  private static void appendPercentEncoded(StringBuilder builder, int codePoint) {
    for (byte value : new String(Character.toChars(codePoint)).getBytes(StandardCharsets.UTF_8)) {
      int unsigned = value & 0xFF;
      builder.append('%').append(HEX_DIGITS[unsigned >>> 4]).append(HEX_DIGITS[unsigned & 0x0F]);
    }
  }

  private static String invalidPathMessage(InvalidPathException exception) {
    return Objects.toString(exception.getReason(), "path is not valid on this runtime");
  }

  private static String stripMailtoPrefix(String email) {
    if (email.regionMatches(true, 0, "mailto:", 0, 7)) {
      return email.substring(7);
    }
    return email;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
