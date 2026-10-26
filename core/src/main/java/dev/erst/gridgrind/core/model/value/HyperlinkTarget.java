package dev.erst.gridgrind.core.model.value;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;

/** Canonical protocol hyperlink target used consistently across request and response payloads. */
@Union(discriminator = "type", catalog = true, displayOrder = 19)
public sealed interface HyperlinkTarget
    permits dev.erst.gridgrind.core.model.value.HyperlinkTarget.Url,
        dev.erst.gridgrind.core.model.value.HyperlinkTarget.Email,
        dev.erst.gridgrind.core.model.value.HyperlinkTarget.File,
        dev.erst.gridgrind.core.model.value.HyperlinkTarget.Document {

  /** Absolute URL hyperlink target such as {@code https://example.com/report}. */
  @Variant(id = "URL", summary = "Attach an absolute URL target.")
  record Url(String target) implements HyperlinkTarget {
    public Url {
      target = ProtocolHyperlinkUrlSupport.normalizeUrlTarget(target);
    }
  }

  /** Email hyperlink target stored without the {@code mailto:} prefix. */
  @Variant(id = "EMAIL", summary = "Attach an email target without the mailto: prefix.")
  record Email(String email) implements HyperlinkTarget {
    public Email {
      email = ProtocolHyperlinkUrlSupport.normalizeEmailTarget(email);
    }
  }

  /**
   * File hyperlink target pointing at a local or shared file-system path.
   *
   * <p>{@code path} accepts either a plain path string or a {@code file:} URI and is normalized to
   * a plain path string for storage and readback.
   */
  @Variant(
      id = "FILE",
      summary =
          "Attach a local or shared file path. Accepts plain paths or file: URIs and normalizes to a plain path string.",
      noteRefs = {"fileHyperlinkRelativePathRule"})
  record File(String path) implements HyperlinkTarget {
    public File {
      path = ProtocolHyperlinkFileSupport.normalizeFileTarget(path);
    }
  }

  /** Internal workbook hyperlink target such as a sheet-and-cell location or defined name. */
  @Variant(id = "DOCUMENT", summary = "Attach an internal workbook target.")
  record Document(String target) implements HyperlinkTarget {
    public Document {
      target = ProtocolHyperlinkTargetSupport.normalizeDocumentTarget(target);
    }
  }
}
