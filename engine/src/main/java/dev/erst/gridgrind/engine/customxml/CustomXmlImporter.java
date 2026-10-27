package dev.erst.gridgrind.engine.customxml;

import java.io.IOException;
import javax.xml.xpath.XPathExpressionException;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.xml.sax.SAXException;

/** Imports XML data into a workbook custom-XML mapping. */
@FunctionalInterface
public interface CustomXmlImporter {
  /** Imports the provided XML payload into one existing workbook mapping. */
  void importXml(XSSFMap mapping, String xml)
      throws IOException, SAXException, XPathExpressionException;
}
