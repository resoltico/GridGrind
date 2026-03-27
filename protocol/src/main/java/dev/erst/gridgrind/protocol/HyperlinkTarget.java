package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Objects;

/** Protocol-facing hyperlink target used by cell hyperlink operations. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = HyperlinkTarget.Url.class, name = "URL"),
  @JsonSubTypes.Type(value = HyperlinkTarget.Email.class, name = "EMAIL"),
  @JsonSubTypes.Type(value = HyperlinkTarget.File.class, name = "FILE"),
  @JsonSubTypes.Type(value = HyperlinkTarget.Document.class, name = "DOCUMENT")
})
public sealed interface HyperlinkTarget
    permits HyperlinkTarget.Url,
        HyperlinkTarget.Email,
        HyperlinkTarget.File,
        HyperlinkTarget.Document {

  /** Converts this protocol hyperlink target into the workbook-core representation. */
  ExcelHyperlink toExcelHyperlink();

  /** Absolute URL hyperlink target such as {@code https://example.com/report}. */
  record Url(String target) implements HyperlinkTarget {
    public Url {
      target = requireNonBlank(target, "target");
      requireAbsoluteUri(target);
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.Url(target);
    }
  }

  /** Email hyperlink target stored without the {@code mailto:} prefix. */
  record Email(String email) implements HyperlinkTarget {
    public Email {
      email = normalizeEmail(email);
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.Email(email);
    }
  }

  /** File hyperlink target pointing at a local or shared file-system path. */
  record File(String target) implements HyperlinkTarget {
    public File {
      target = requireNonBlank(target, "target");
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.File(target);
    }
  }

  /** Internal workbook hyperlink target such as a sheet-and-cell location or defined name. */
  record Document(String target) implements HyperlinkTarget {
    public Document {
      target = requireNonBlank(target, "target");
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.Document(target);
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
    String normalized = requireNonBlank(email, "email");
    if (normalized.regionMatches(true, 0, "mailto:", 0, 7)) {
      normalized = normalized.substring(7);
    }
    if (normalized.isBlank()) {
      throw new IllegalArgumentException("email must not be blank");
    }
    return normalized;
  }
}
