package dev.erst.gridgrind.excel;

import java.io.OutputStream;
import javax.xml.transform.TransformerException;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.xml.sax.SAXException;

/** Seams custom-XML export so controller error handling is testable. */
@FunctionalInterface
interface ExcelCustomXmlExporter {
  /** Writes one workbook mapping as XML using the requested encoding and validation mode. */
  void export(XSSFMap mapping, OutputStream output, String encoding, boolean validateSchema)
      throws TransformerException, SAXException;
}
