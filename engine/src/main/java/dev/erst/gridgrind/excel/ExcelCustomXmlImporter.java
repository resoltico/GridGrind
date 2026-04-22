package dev.erst.gridgrind.excel;

import java.io.IOException;
import javax.xml.xpath.XPathExpressionException;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.xml.sax.SAXException;

/** Seams custom-XML import so controller error handling is testable. */
@FunctionalInterface
interface ExcelCustomXmlImporter {
  /** Imports the provided XML payload into one existing workbook mapping. */
  void importXml(XSSFMap mapping, String xml)
      throws IOException, SAXException, XPathExpressionException;
}
