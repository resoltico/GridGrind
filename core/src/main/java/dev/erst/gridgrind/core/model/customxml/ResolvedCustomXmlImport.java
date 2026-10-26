package dev.erst.gridgrind.core.model.customxml;

import java.util.Objects;

/** Materialized XML content bound to an existing custom-XML mapping locator. */
public record ResolvedCustomXmlImport(CustomXmlMappingLocator locator, String xml) {
  public ResolvedCustomXmlImport {
    Objects.requireNonNull(locator, "locator must not be null");
    Objects.requireNonNull(xml, "xml must not be null");
    if (xml.isBlank()) {
      throw new IllegalArgumentException("xml must not be blank");
    }
  }
}
