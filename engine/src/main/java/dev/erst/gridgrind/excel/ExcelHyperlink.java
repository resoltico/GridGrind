package dev.erst.gridgrind.excel;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Objects;

/** Immutable workbook-core hyperlink target used for cell hyperlink authoring and analysis. */
public sealed interface ExcelHyperlink
    permits ExcelHyperlink.Url, ExcelHyperlink.Email, ExcelHyperlink.File, ExcelHyperlink.Document {

  /** Returns the stable hyperlink type label. */
  ExcelHyperlinkType type();

  /** Returns the normalized target string for this hyperlink. */
  String target();

  /** External URL hyperlink target. */
  record Url(String target) implements ExcelHyperlink {
    public Url {
      target = requireNonBlank(target, "target");
      requireAbsoluteUri(target);
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.URL;
    }
  }

  /** Email hyperlink target stored without the {@code mailto:} prefix. */
  record Email(String target) implements ExcelHyperlink {
    public Email {
      target = normalizeEmail(target);
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.EMAIL;
    }
  }

  /** File hyperlink target. */
  record File(String target) implements ExcelHyperlink {
    public File {
      target = requireNonBlank(target, "target");
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.FILE;
    }
  }

  /** Internal workbook hyperlink target such as a sheet location or defined name. */
  record Document(String target) implements ExcelHyperlink {
    public Document {
      target = requireNonBlank(target, "target");
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.DOCUMENT;
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static void requireAbsoluteUri(String target) {
    try {
      URI uri = new URI(target);
      if (!uri.isAbsolute()) {
        throw new IllegalArgumentException("target must be an absolute URI with a scheme");
      }
    } catch (URISyntaxException exception) {
      throw new IllegalArgumentException("target must be a valid absolute URI", exception);
    }
  }

  private static String normalizeEmail(String email) {
    String normalized = requireNonBlank(email, "target");
    if (normalized.regionMatches(true, 0, "mailto:", 0, 7)) {
      normalized = normalized.substring(7);
    }
    if (normalized.isBlank()) {
      throw new IllegalArgumentException("target must not be blank");
    }
    return normalized;
  }
}
