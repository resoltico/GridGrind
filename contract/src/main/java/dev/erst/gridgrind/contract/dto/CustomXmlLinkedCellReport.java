package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Factual single-cell mapping entry linked to one custom-XML map. */
public record CustomXmlLinkedCellReport(
    String sheetName, String address, String xpath, String xmlDataType) {
  public CustomXmlLinkedCellReport {
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
