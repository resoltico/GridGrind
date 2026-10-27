package dev.erst.gridgrind.engine.customxml;

import java.io.OutputStream;
import javax.xml.transform.TransformerException;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.xml.sax.SAXException;

/** Exports XML data from a workbook custom-XML mapping. */
@FunctionalInterface
public interface CustomXmlExporter {
  /** Writes one workbook mapping as XML using the requested encoding and validation mode. */
  void export(XSSFMap mapping, OutputStream output, String encoding, boolean validateSchema)
      throws TransformerException, SAXException;
}
