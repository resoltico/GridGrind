package dev.erst.gridgrind.excel;

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
      if (!ExcelHyperlinkValidation.isValidUrlTarget(target)) {
        String scheme = ExcelHyperlinkValidation.absoluteUriScheme(target).orElse(null);
        if ("file".equalsIgnoreCase(scheme)) {
          throw new IllegalArgumentException(
              "target uses file: scheme; use FILE hyperlinks instead");
        }
        if ("mailto".equalsIgnoreCase(scheme)) {
          throw new IllegalArgumentException(
              "target uses mailto: scheme; use EMAIL hyperlinks instead");
        }
        throw new IllegalArgumentException("target must be an absolute URI with a scheme");
      }
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.URL;
    }
  }

  /** Email hyperlink target stored without the {@code mailto:} prefix. */
  record Email(String target) implements ExcelHyperlink {
    public Email {
      target = normalizeEmailTarget(target);
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.EMAIL;
    }
  }

  /** File hyperlink target. */
  record File(String path) implements ExcelHyperlink {
    public File {
      path = ExcelHyperlinkValidation.normalizeFileTarget(path);
    }

    @Override
    public ExcelHyperlinkType type() {
      return ExcelHyperlinkType.FILE;
    }

    @Override
    public String target() {
      return path;
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

  private static String normalizeEmailTarget(String email) {
    String normalized =
        ExcelHyperlinkValidation.stripMailtoPrefix(requireNonBlank(email, "target"));
    if (!ExcelHyperlinkValidation.isValidEmailTarget(normalized)) {
      throw new IllegalArgumentException("target must be an email address");
    }
    return normalized;
  }
}
