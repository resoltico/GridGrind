package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.time.Instant;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Regression coverage for product-owned workbook document metadata normalization. */
class ExcelWorkbookDocumentMetadataSupportTest {
  @Test
  void normalizeForSaveAssignsProductOwnedMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");

      ExcelWorkbookDocumentMetadataSupport.normalizeForSave(workbook);

      var properties = workbook.getProperties();
      var core = properties.getCoreProperties();
      var extended = properties.getExtendedProperties();

      assertEquals("GridGrind", core.getCreator());
      assertEquals("GridGrind", core.getLastModifiedByUser());
      assertEquals("GridGrind", extended.getApplication());
      assertNull(core.getLastPrinted());

      String sourceDateEpoch = System.getenv("SOURCE_DATE_EPOCH");
      if (sourceDateEpoch == null || sourceDateEpoch.isBlank()) {
        Instant expected = Instant.parse("2000-01-01T00:00:00Z");
        assertNotNull(core.getCreated());
        assertNotNull(core.getModified());
        assertEquals(expected, core.getCreated().toInstant());
        assertEquals(expected, core.getModified().toInstant());
      } else {
        Instant expected = Instant.ofEpochSecond(Long.parseLong(sourceDateEpoch.trim()));
        assertNotNull(core.getCreated());
        assertNotNull(core.getModified());
        assertEquals(expected, core.getCreated().toInstant());
        assertEquals(expected, core.getModified().toInstant());
      }
      assertNotNull(workbook.getSheet("Budget"));
    }
  }

  @Test
  void deterministicTimestampParserCoversBlankValidAndRejectedValues() {
    assertEquals(
        Instant.parse("2000-01-01T00:00:00Z"),
        ExcelWorkbookDocumentMetadataSupport.deterministicTimestamp(null)
            .orElseThrow()
            .toInstant());
    assertEquals(
        Instant.parse("2000-01-01T00:00:00Z"),
        ExcelWorkbookDocumentMetadataSupport.deterministicTimestamp(" ").orElseThrow().toInstant());
    assertEquals(
        Instant.ofEpochSecond(1_717_782_600L),
        ExcelWorkbookDocumentMetadataSupport.deterministicTimestamp("1717782600")
            .orElseThrow()
            .toInstant());
    assertEquals(
        "SOURCE_DATE_EPOCH must be >= 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> ExcelWorkbookDocumentMetadataSupport.deterministicTimestamp("-1"))
            .getMessage());
    assertEquals(
        "SOURCE_DATE_EPOCH must be one integer Unix timestamp in whole seconds",
        assertThrows(
                IllegalArgumentException.class,
                () -> ExcelWorkbookDocumentMetadataSupport.deterministicTimestamp("not-a-number"))
            .getMessage());
  }
}
