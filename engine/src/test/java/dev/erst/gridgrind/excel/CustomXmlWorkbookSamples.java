package dev.erst.gridgrind.excel;

import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Upstream-backed custom-XML workbook fixtures for deterministic engine tests. */
final class CustomXmlWorkbookSamples {
  private static final String SIMPLE_CUSTOM_XML_WORKBOOK_RESOURCE = "CustomXMLMappings.xlsx";

  private CustomXmlWorkbookSamples() {}

  static ExcelWorkbook openSimpleCustomXmlWorkbook() throws IOException {
    try (InputStream input =
        CustomXmlWorkbookSamples.class.getResourceAsStream(SIMPLE_CUSTOM_XML_WORKBOOK_RESOURCE)) {
      if (input == null) {
        throw new IllegalStateException(
            "Missing test workbook resource " + SIMPLE_CUSTOM_XML_WORKBOOK_RESOURCE);
      }
      return ExcelWorkbook.wrap(new XSSFWorkbook(input));
    }
  }
}
