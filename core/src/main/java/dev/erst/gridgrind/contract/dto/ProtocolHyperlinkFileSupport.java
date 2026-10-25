package dev.erst.gridgrind.contract.dto;

import java.net.URI;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.Objects;
import java.util.regex.Pattern;

/** Protocol-owned file-hyperlink normalization and POI-address rendering helpers. */
final class ProtocolHyperlinkFileSupport {
  private static final char[] HEX_DIGITS = "0123456789ABCDEF".toCharArray();
  private static final String SAFE_URI_PATH_CHARACTERS =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~/:";
  private static final Pattern WINDOWS_DRIVE_PATH = Pattern.compile("^[A-Za-z]:[/\\\\].*");

  private ProtocolHyperlinkFileSupport() {}

  static String normalizeFileTarget(String path) {
    Objects.requireNonNull(path, "path must not be null");
    if (path.isBlank()) {
      throw new IllegalArgumentException("path must not be blank");
    }
    if (looksLikeFileUri(path)) {
      return normalizeFileUri(path);
    }
    if (ProtocolHyperlinkUrlSupport.looksLikeAbsoluteUri(path)) {
      throw new IllegalArgumentException("path must be a local file path or file: URI");
    }
    return decodeEscapedRelativePath(path);
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

  private static boolean looksLikeWindowsDrivePath(String path) {
    return WINDOWS_DRIVE_PATH.matcher(path).matches();
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
}
