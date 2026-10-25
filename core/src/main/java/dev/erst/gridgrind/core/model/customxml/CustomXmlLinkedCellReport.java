package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Factual single-cell mapping entry linked to one custom-XML map. */
public record CustomXmlLinkedCellReport(
    SheetLabel sheetName, String address, String xpath, String xmlDataType) {
  public CustomXmlLinkedCellReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    address = requireNonBlank(address, "address");
    xpath = requireNonBlank(xpath, "xpath");
    xmlDataType = requireNonBlank(xmlDataType, "xmlDataType");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
