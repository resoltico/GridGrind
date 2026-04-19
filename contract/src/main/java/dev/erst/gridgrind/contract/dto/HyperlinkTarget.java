package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

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

  /** Absolute URL hyperlink target such as {@code https://example.com/report}. */
  record Url(String target) implements HyperlinkTarget {
    public Url {
      target = ProtocolHyperlinkSupport.normalizeUrlTarget(target);
    }
  }

  /** Email hyperlink target stored without the {@code mailto:} prefix. */
  record Email(String email) implements HyperlinkTarget {
    public Email {
      email = ProtocolHyperlinkSupport.normalizeEmailTarget(email);
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
      path = ProtocolHyperlinkSupport.normalizeFileTarget(path);
    }
  }

  /** Internal workbook hyperlink target such as a sheet-and-cell location or defined name. */
  record Document(String target) implements HyperlinkTarget {
    public Document {
      target = ProtocolHyperlinkSupport.normalizeDocumentTarget(target);
    }
  }
}
