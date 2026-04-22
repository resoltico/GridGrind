package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Authoritative workbook-level XML import payload bound to one existing custom-XML mapping. */
public record CustomXmlImportInput(CustomXmlMappingLocator locator, TextSourceInput xml) {
  public CustomXmlImportInput {
    Objects.requireNonNull(locator, "locator must not be null");
    Objects.requireNonNull(xml, "xml must not be null");
    if (xml instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException("xml must not be blank");
    }
  }
}
