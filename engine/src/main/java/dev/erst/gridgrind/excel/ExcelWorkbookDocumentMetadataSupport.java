package dev.erst.gridgrind.excel;

import java.time.Instant;
import java.util.Date;
import java.util.Optional;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Owns deterministic, product-owned OOXML workbook document metadata at save time. */
final class ExcelWorkbookDocumentMetadataSupport {
  private static final String PRODUCT_NAME = "GridGrind";

  private ExcelWorkbookDocumentMetadataSupport() {}

  static void normalizeForSave(XSSFWorkbook workbook) {
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
    return deterministicTimestamp(System.getenv("SOURCE_DATE_EPOCH"));
  }

  static Optional<Date> deterministicTimestamp(String sourceDateEpoch) {
    if (sourceDateEpoch == null || sourceDateEpoch.isBlank()) {
      return Optional.empty();
    }
    try {
      long epochSeconds = Long.parseLong(sourceDateEpoch.trim());
      if (epochSeconds < 0) {
        throw new IllegalArgumentException("SOURCE_DATE_EPOCH must be >= 0");
      }
      return Optional.of(Date.from(Instant.ofEpochSecond(epochSeconds)));
    } catch (NumberFormatException exception) {
      throw new IllegalArgumentException(
          "SOURCE_DATE_EPOCH must be one integer Unix timestamp in whole seconds", exception);
    }
  }
}
