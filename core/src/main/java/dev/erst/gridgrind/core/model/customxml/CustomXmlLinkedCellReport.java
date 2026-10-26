package dev.erst.gridgrind.core.model.customxml;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Factual single-cell mapping entry linked to one custom-XML map. */
@CatalogType(
    id = "CustomXmlLinkedCellReport",
    summary = "One single-cell binding linked to a custom-XML mapping.",
    displayOrder = 105)
public record CustomXmlLinkedCellReport(
    SheetLabel sheetName, CellAddress address, String xpath, String xmlDataType) {
  public CustomXmlLinkedCellReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    java.util.Objects.requireNonNull(address, "address must not be null");
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
