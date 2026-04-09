package dev.erst.gridgrind.excel;

import java.net.URI;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.Objects;

/** Normalizes and resolves local file hyperlink paths independently of Apache POI quirks. */
final class ExcelFileHyperlinkTargets {
  private static final char[] HEX_DIGITS = "0123456789ABCDEF".toCharArray();

  private ExcelFileHyperlinkTargets() {}

  /**
   * Normalizes one request or workbook file hyperlink value to the canonical stored path string.
   */
  static String normalizePath(String path) {
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

  /**
   * Converts one request-side file hyperlink path into the escaped address string required by POI.
   */
  static String toPoiAddress(String path) {
    String normalizedPath = normalizePath(path);
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

  /**
   * Resolves one normalized file hyperlink path against the workbook's current filesystem anchor.
   */
  static FileHyperlinkResolution resolve(String path, WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    String normalizedPath;
    try {
      normalizedPath = normalizePath(path);
    } catch (IllegalArgumentException exception) {
      return new FileHyperlinkResolution.MalformedPath(path, exception.getMessage());
    }

    Path candidate;
    try {
      candidate = Path.of(normalizedPath);
    } catch (InvalidPathException exception) {
      return new FileHyperlinkResolution.MalformedPath(
          normalizedPath, invalidPathMessage(exception));
    }

    if (candidate.isAbsolute()) {
      return new FileHyperlinkResolution.ResolvedPath(normalizedPath, candidate);
    }
    return switch (workbookLocation) {
      case WorkbookLocation.StoredWorkbook storedWorkbook ->
          new FileHyperlinkResolution.ResolvedPath(
              normalizedPath, storedWorkbook.baseDirectory().orElseThrow().resolve(candidate));
      case WorkbookLocation.UnsavedWorkbook _ ->
          new FileHyperlinkResolution.UnresolvedRelativePath(normalizedPath);
    };
  }

  /** Returns whether the normalized stored path is relative in GridGrind's file-target contract. */
  static boolean isRelativeStoredPath(String path) {
    Objects.requireNonNull(path, "path must not be null");
    if (looksLikeWindowsDrivePath(path)) {
      return false;
    }
    try {
      return !Path.of(path).isAbsolute();
    } catch (InvalidPathException exception) {
      return false;
    }
  }

  private static boolean looksLikeFileUri(String path) {
    return path.regionMatches(true, 0, "file:", 0, 5);
  }

  private static String normalizeFileUri(String path) {
    try {
      URI uri = new URI(path);
      try {
        return Path.of(uri).toString();
      } catch (RuntimeException exception) {
        String authority = uri.getAuthority();
        String uriPath = uri.getPath();
        if (uriPath == null || uriPath.isBlank()) {
          throw new IllegalArgumentException("path must contain a file-system path", exception);
        }
        return authority == null ? uriPath : "//" + authority + uriPath;
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
    return path.length() >= 3
        && Character.isLetter(path.charAt(0))
        && path.charAt(1) == ':'
        && (path.charAt(2) == '/' || path.charAt(2) == '\\');
  }

  private static String decodeEscapedRelativePath(String path) {
    if (!path.contains("%")) {
      return path;
    }
    try {
      URI uri = new URI(path);
      if (uri.isAbsolute()) {
        return path;
      }
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
    if (codePoint >= 'A' && codePoint <= 'Z') {
      return true;
    }
    if (codePoint >= 'a' && codePoint <= 'z') {
      return true;
    }
    if (codePoint >= '0' && codePoint <= '9') {
      return true;
    }
    return switch (codePoint) {
      case '-', '.', '_', '~', '/', ':' -> true;
      default -> false;
    };
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
