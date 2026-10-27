package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Materialized XML content bound to an existing custom-XML mapping locator. */
public record ResolvedCustomXmlImport(CustomXmlMappingLocator locator, String xml) {
  public ResolvedCustomXmlImport {
    Objects.requireNonNull(locator, "locator must not be null");
    xml = TextValidation.requireNonBlank(xml, "xml");
  }
}
