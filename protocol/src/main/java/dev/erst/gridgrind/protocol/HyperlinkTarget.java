package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelHyperlink;

/** Canonical protocol hyperlink target used consistently across request and response payloads. */
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
      target = new ExcelHyperlink.Url(target).target();
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.Url(target);
    }
  }

  /** Email hyperlink target stored without the {@code mailto:} prefix. */
  record Email(String email) implements HyperlinkTarget {
    public Email {
      email = new ExcelHyperlink.Email(email).target();
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.Email(email);
    }
  }

  /**
   * File hyperlink target pointing at a local or shared file-system path.
   *
   * <p>{@code path} accepts either a plain path string or a {@code file:} URI and is normalized to
   * a plain path string for storage and readback.
   */
  record File(String path) implements HyperlinkTarget {
    public File {
      path = new ExcelHyperlink.File(path).target();
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.File(path);
    }
  }

  /** Internal workbook hyperlink target such as a sheet-and-cell location or defined name. */
  record Document(String target) implements HyperlinkTarget {
    public Document {
      target = new ExcelHyperlink.Document(target).target();
    }

    @Override
    public ExcelHyperlink toExcelHyperlink() {
      return new ExcelHyperlink.Document(target);
    }
  }
}
