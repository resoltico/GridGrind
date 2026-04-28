package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.lang.invoke.MethodHandles;
import java.util.List;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for POI clone-preparation support on hyperlink-bearing sheets. */
class ExcelSheetClonePreparationSupportTest {
  private final ExcelSheetClonePreparationSupport support = new ExcelSheetClonePreparationSupport();

  @Test
  void prepareSourceSheetForCloneMaterializesUnsavedExternalHyperlinksIdempotently()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Source");
      sheet.createRow(1).createCell(0).setCellValue("Link");
      XSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
      hyperlink.setAddress("https://example.com/report");
      sheet.getRow(1).getCell(0).setHyperlink(hyperlink);

      XSSFHyperlink unsavedHyperlink = sheet.getHyperlinkList().getFirst();
      assertNull(unsavedHyperlink.getCTHyperlink().getId());
      assertEquals(
          0,
          sheet
              .getPackagePart()
              .getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation())
              .size());

      support.prepareSourceSheetForClone(sheet);

      XSSFHyperlink prepared = sheet.getHyperlinkList().getFirst();
      String relationId = prepared.getCTHyperlink().getId();
      assertNotNull(relationId);
      assertEquals(
          "https://example.com/report",
          sheet.getPackagePart().getRelationship(relationId).getTargetURI().toString());
      assertEquals(
          1,
          sheet
              .getPackagePart()
              .getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation())
              .size());

      support.prepareSourceSheetForClone(sheet);

      XSSFHyperlink preparedAgain = sheet.getHyperlinkList().getFirst();
      assertEquals(relationId, preparedAgain.getCTHyperlink().getId());
      assertEquals(
          1,
          sheet
              .getPackagePart()
              .getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation())
              .size());
    }
  }

  @Test
  void prepareSourceSheetForCloneLeavesDocumentHyperlinksWithoutExternalRelationships()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Source");
      sheet.createRow(0).createCell(0).setCellValue("Jump");
      XSSFHyperlink hyperlink =
          workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
      hyperlink.setAddress("Source!A1");
      sheet.getRow(0).getCell(0).setHyperlink(hyperlink);

      support.prepareSourceSheetForClone(sheet);

      assertNull(sheet.getHyperlinkList().getFirst().getCTHyperlink().getId());
      assertEquals(
          0,
          sheet
              .getPackagePart()
              .getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation())
              .size());
    }
  }

  @Test
  void prepareSourceSheetForCloneIgnoresNullHyperlinkEntries() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Source");
      sheet.createRow(0).createCell(0).setCellValue("Link");
      XSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
      hyperlink.setAddress("https://example.com/null-guard");
      sheet.getRow(0).getCell(0).setHyperlink(hyperlink);

      @SuppressWarnings("unchecked")
      List<XSSFHyperlink> hyperlinks =
          (List<XSSFHyperlink>)
              ExcelSheetClonePreparationSupport.requireHyperlinksField(MethodHandles.lookup())
                  .get(sheet);
      hyperlinks.add(null);

      assertDoesNotThrow(() -> support.prepareSourceSheetForClone(sheet));
      assertNull(hyperlinks.getLast());
    }
  }

  @Test
  void reflectivePoiAccessorsRejectInvalidLookups() {
    NullPointerException hyperlinksFieldException =
        assertThrows(
            NullPointerException.class,
            () -> ExcelSheetClonePreparationSupport.requireHyperlinksField(null));
    assertEquals("lookup must not be null", hyperlinksFieldException.getMessage());

    NullPointerException hyperlinkConstructorException =
        assertThrows(
            NullPointerException.class,
            () -> ExcelSheetClonePreparationSupport.requireHyperlinkConstructor(null));
    assertEquals("lookup must not be null", hyperlinkConstructorException.getMessage());

    assertDoesNotThrow(
        () -> ExcelSheetClonePreparationSupport.requireHyperlinksField(MethodHandles.lookup()));
    assertDoesNotThrow(
        () ->
            ExcelSheetClonePreparationSupport.requireHyperlinkConstructor(MethodHandles.lookup()));

    IllegalStateException hyperlinksFieldLookupFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelSheetClonePreparationSupport.requireHyperlinksField(
                    MethodHandles.publicLookup()));
    assertTrue(
        hyperlinksFieldLookupFailure
            .getMessage()
            .contains("Apache POI private contract unavailable"));
    assertTrue(
        hyperlinksFieldLookupFailure
            .getMessage()
            .contains(
                ExcelSheetClonePreparationSupport.HYPERLINKS_FIELD_CONTRACT.affectedSurface()));

    IllegalStateException hyperlinkConstructorLookupFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                ExcelSheetClonePreparationSupport.requireHyperlinkConstructor(
                    MethodHandles.publicLookup()));
    assertTrue(
        hyperlinkConstructorLookupFailure
            .getMessage()
            .contains(
                ExcelSheetClonePreparationSupport.HYPERLINK_CONSTRUCTOR_CONTRACT
                    .affectedSurface()));
  }
}
