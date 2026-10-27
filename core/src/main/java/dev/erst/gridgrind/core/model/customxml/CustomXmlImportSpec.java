package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Objects;

/** Authoritative workbook-level XML import payload bound to one existing custom-XML mapping. */
@CatalogType(
    id = "CustomXmlImportSpec",
    summary =
        "One custom-XML import payload targeting an existing workbook mapping plus the XML content to import.",
    displayOrder = 32)
public record CustomXmlImportSpec(CustomXmlMappingLocator locator, TextSourceSpec xml) {
  public CustomXmlImportSpec {
    Objects.requireNonNull(locator, "locator must not be null");
    Objects.requireNonNull(xml, "xml must not be null");
    if (xml instanceof TextSourceSpec.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException("xml must not be blank");
    }
  }
}
