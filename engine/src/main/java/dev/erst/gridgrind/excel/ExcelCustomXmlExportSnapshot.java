package dev.erst.gridgrind.excel;

import java.util.Objects;

/** XML document exported from one workbook custom-XML mapping. */
public record ExcelCustomXmlExportSnapshot(
    ExcelCustomXmlMappingSnapshot mapping, String encoding, boolean schemaValidated, String xml) {
  public ExcelCustomXmlExportSnapshot {
    Objects.requireNonNull(mapping, "mapping must not be null");
    encoding = requireNonBlank(encoding, "encoding");
    xml = requireNonBlank(xml, "xml");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
