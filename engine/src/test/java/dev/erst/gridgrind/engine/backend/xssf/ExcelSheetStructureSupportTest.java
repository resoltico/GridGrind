package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
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

      assertFalse(SheetStructure.hasHeaderValue(null));
      assertFalse(SheetStructure.hasHeaderValue(row.getCell(0)));
      assertFalse(SheetStructure.hasHeaderValue(row.getCell(1)));
      assertFalse(SheetStructure.hasHeaderValue(cellStub(CellType._NONE, "", "", "")));
      assertFalse(SheetStructure.hasHeaderValue(cellStub(CellType.FORMULA, "", "", "")));
      assertTrue(SheetStructure.hasHeaderValue(row.getCell(3)));
      assertTrue(SheetStructure.hasHeaderValue(row.getCell(4)));
      assertTrue(SheetStructure.hasHeaderValue(row.getCell(5)));
      assertTrue(SheetStructure.hasHeaderValue(row.getCell(6)));
    }
  }

  @Test
  void headerRowMissingReflectsActualHeaderContent() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      sheet.createRow(1).createCell(0).setCellValue("Ada");

      assertTrue(SheetStructure.headerRowMissing(sheet, CellBounds.parse("A1:B2")));

      var headerRow = sheet.createRow(0);
      headerRow.createCell(0).setCellValue("Owner");

      assertFalse(SheetStructure.headerRowMissing(sheet, CellBounds.parse("A1:B2")));
    }
  }

  @Test
  void parseAndIntersectHandleInvalidAndBoundaryRanges() {
    assertTrue(SheetStructure.parseOptionalRange("A0:B2").isEmpty());
    assertEquals("A1:B2", SheetStructure.formatRange(CellBounds.parse("A1:B2")));

    CellBounds first = CellBounds.parse("A1:B2");
    CellBounds beforeFirst = CellBounds.parse("A5:B6");
    CellBounds touching = CellBounds.parse("B2:C3");
    CellBounds noColumnOverlap = CellBounds.parse("C1:D2");
    CellBounds noRowOverlap = CellBounds.parse("A3:B4");

    assertFalse(SheetStructure.intersects(beforeFirst, first));
    assertTrue(SheetStructure.intersects(first, touching));
    assertFalse(SheetStructure.intersects(first, noColumnOverlap));
    assertFalse(SheetStructure.intersects(first, noRowOverlap));
  }

  @Test
  void defaultConstructorAndPreviewHelpersCoverBlankAndAnnotatedCells() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Ops");
      assertEquals(0, sheet.getPhysicalNumberOfRows());
      assertFalse(SheetStructure.shouldPreview((org.apache.poi.ss.usermodel.Cell) null));

      var row = sheet.createRow(0);
      var blankCell = row.createCell(0);
      assertFalse(SheetStructure.shouldPreview(blankCell));

      var hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
      hyperlink.setAddress("https://example.com");
      blankCell.setHyperlink(hyperlink);

      assertTrue(SheetStructure.shouldPreview(blankCell));

      var commentCell = row.createCell(1);
      var anchor = workbook.getCreationHelper().createClientAnchor();
      var comment = sheet.createDrawingPatriarch().createCellComment(anchor);
      comment.setString(workbook.getCreationHelper().createRichTextString("Review"));
      commentCell.setCellComment(comment);

      assertTrue(SheetStructure.shouldPreview(commentCell));
      assertFalse(SheetStructure.shouldPreview(CellType.BLANK, (short) 0, false, false));
      assertTrue(SheetStructure.shouldPreview(CellType.BLANK, (short) 1, false, false));
      assertTrue(SheetStructure.shouldPreview(CellType.BLANK, (short) 0, true, false));
      assertTrue(SheetStructure.shouldPreview(CellType.BLANK, (short) 0, false, true));
      assertTrue(SheetStructure.shouldPreview(CellType.STRING, (short) 0, false, false));
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
