package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual single-cell mapping entry linked to one custom-XML map. */
public record ExcelCustomXmlLinkedCellSnapshot(
    String sheetName, String address, String xpath, String xmlDataType) {
  public ExcelCustomXmlLinkedCellSnapshot {
    sheetName = requireNonBlank(sheetName, "sheetName");
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
