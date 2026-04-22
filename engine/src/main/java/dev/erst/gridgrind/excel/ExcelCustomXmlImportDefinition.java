package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authoritative workbook-level XML import payload bound to one existing custom-XML mapping. */
public record ExcelCustomXmlImportDefinition(ExcelCustomXmlMappingLocator mapping, String xml) {
  public ExcelCustomXmlImportDefinition {
    Objects.requireNonNull(mapping, "mapping must not be null");
    Objects.requireNonNull(xml, "xml must not be null");
    if (xml.isBlank()) {
      throw new IllegalArgumentException("xml must not be blank");
    }
  }
}
