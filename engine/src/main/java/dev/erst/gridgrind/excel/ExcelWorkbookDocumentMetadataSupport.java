package dev.erst.gridgrind.excel;

import java.util.Date;
import java.util.Optional;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Owns deterministic, product-owned OOXML workbook document metadata at save time. */
public final class ExcelWorkbookDocumentMetadataSupport {
  private static final String PRODUCT_NAME = "GridGrind";

  private ExcelWorkbookDocumentMetadataSupport() {}

  /** Applies deterministic product-owned OOXML core and extended properties before saving. */
  public static void normalizeForSave(XSSFWorkbook workbook) {
    java.util.Objects.requireNonNull(workbook, "workbook must not be null");
    POIXMLProperties properties = workbook.getProperties();
    POIXMLProperties.CoreProperties core = properties.getCoreProperties();
    POIXMLProperties.ExtendedProperties extended = properties.getExtendedProperties();
    Optional<Date> deterministicTimestamp = deterministicTimestamp();

    core.setCreator(PRODUCT_NAME);
    core.setLastModifiedByUser(PRODUCT_NAME);
    core.setCreated(deterministicTimestamp);
    core.setModified(deterministicTimestamp);
    core.setLastPrinted(Optional.empty());
    extended.setApplication(PRODUCT_NAME);
  }

  private static Optional<Date> deterministicTimestamp() {
    return ExcelDeterministicWorkbookArtifactSupport.deterministicTimestamp();
  }

  static Optional<Date> deterministicTimestamp(String sourceDateEpoch) {
    return ExcelDeterministicWorkbookArtifactSupport.deterministicTimestamp(sourceDateEpoch);
  }
}
