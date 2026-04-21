package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.lang.reflect.Proxy;
import java.util.Objects;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for shared sheet-structure helpers used by table and autofilter controllers. */
class ExcelSheetStructureSupportTest {
  @Test
  void hasHeaderValueRecognizesSupportedCellKinds() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var row = sheet.createRow(0);
      row.createCell(0).setBlank();
      row.createCell(1).setCellValue("   ");
      row.createCell(3).setCellValue("Owner");
      row.createCell(4).setCellFormula("A1");
      row.createCell(5).setCellValue(12.0);
      row.createCell(6).setCellValue(true);

      assertFalse(ExcelSheetStructureSupport.hasHeaderValue(null));
      assertFalse(ExcelSheetStructureSupport.hasHeaderValue(row.getCell(0)));
      assertFalse(ExcelSheetStructureSupport.hasHeaderValue(row.getCell(1)));
      assertFalse(ExcelSheetStructureSupport.hasHeaderValue(cellStub(CellType._NONE, "", "", "")));
      assertFalse(
          ExcelSheetStructureSupport.hasHeaderValue(cellStub(CellType.FORMULA, "", "", "")));
      assertTrue(ExcelSheetStructureSupport.hasHeaderValue(row.getCell(3)));
      assertTrue(ExcelSheetStructureSupport.hasHeaderValue(row.getCell(4)));
      assertTrue(ExcelSheetStructureSupport.hasHeaderValue(row.getCell(5)));
      assertTrue(ExcelSheetStructureSupport.hasHeaderValue(row.getCell(6)));
    }
  }

  @Test
  void headerRowMissingReflectsActualHeaderContent() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      sheet.createRow(1).createCell(0).setCellValue("Ada");

      assertTrue(ExcelSheetStructureSupport.headerRowMissing(sheet, ExcelRange.parse("A1:B2")));

      var headerRow = sheet.createRow(0);
      headerRow.createCell(0).setCellValue("Owner");

      assertFalse(ExcelSheetStructureSupport.headerRowMissing(sheet, ExcelRange.parse("A1:B2")));
    }
  }

  @Test
  void parseAndIntersectHandleInvalidAndBoundaryRanges() {
    assertNull(ExcelSheetStructureSupport.parseRangeOrNull("A0:B2"));
    assertEquals("A1:B2", ExcelSheetStructureSupport.formatRange(ExcelRange.parse("A1:B2")));

    ExcelRange first = ExcelRange.parse("A1:B2");
    ExcelRange beforeFirst = ExcelRange.parse("A5:B6");
    ExcelRange touching = ExcelRange.parse("B2:C3");
    ExcelRange noColumnOverlap = ExcelRange.parse("C1:D2");
    ExcelRange noRowOverlap = ExcelRange.parse("A3:B4");

    assertFalse(ExcelSheetStructureSupport.intersects(beforeFirst, first));
    assertTrue(ExcelSheetStructureSupport.intersects(first, touching));
    assertFalse(ExcelSheetStructureSupport.intersects(first, noColumnOverlap));
    assertFalse(ExcelSheetStructureSupport.intersects(first, noRowOverlap));
  }

  @Test
  void defaultConstructorAndPreviewHelpersCoverBlankAndAnnotatedCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      var support =
          new ExcelSheetStructureSupport(
              sheet,
              ExcelFormulaRuntime.poi(workbook.getCreationHelper().createFormulaEvaluator()));

      assertEquals(0, support.physicalRowCount());
      assertFalse(
          ExcelSheetStructureSupport.shouldPreview((org.apache.poi.ss.usermodel.Cell) null));

      var row = sheet.createRow(0);
      var blankCell = row.createCell(0);
      assertFalse(ExcelSheetStructureSupport.shouldPreview(blankCell));

      var hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
      hyperlink.setAddress("https://example.com");
      blankCell.setHyperlink(hyperlink);

      assertTrue(ExcelSheetStructureSupport.shouldPreview(blankCell));

      var commentCell = row.createCell(1);
      var anchor = workbook.getCreationHelper().createClientAnchor();
      var comment = sheet.createDrawingPatriarch().createCellComment(anchor);
      comment.setString(workbook.getCreationHelper().createRichTextString("Review"));
      commentCell.setCellComment(comment);

      assertTrue(ExcelSheetStructureSupport.shouldPreview(commentCell));
      assertFalse(
          ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 0, false, false));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 1, false, false));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 0, true, false));
      assertTrue(ExcelSheetStructureSupport.shouldPreview(CellType.BLANK, (short) 0, false, true));
      assertTrue(
          ExcelSheetStructureSupport.shouldPreview(CellType.STRING, (short) 0, false, false));
    }
  }

  private static org.apache.poi.ss.usermodel.Cell cellStub(
      CellType cellType, String stringValue, String formulaValue, String displayValue) {
    return (org.apache.poi.ss.usermodel.Cell)
        Proxy.newProxyInstance(
            Objects.requireNonNull(Thread.currentThread().getContextClassLoader()),
            new Class<?>[] {org.apache.poi.ss.usermodel.Cell.class},
            (proxy, method, args) ->
                switch (method.getName()) {
                  case "getCellType" -> cellType;
                  case "getStringCellValue" -> stringValue;
                  case "getCellFormula" -> formulaValue;
                  case "toString" -> displayValue;
                  case "hashCode" -> System.identityHashCode(proxy);
                  case "equals" -> false;
                  default -> throw new UnsupportedOperationException(method.getName());
                });
  }
}
