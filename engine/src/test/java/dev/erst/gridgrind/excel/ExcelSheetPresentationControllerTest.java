package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Direct tests for authoritative sheet-presentation read and write behavior. */
class ExcelSheetPresentationControllerTest {
  @Test
  void setAndClearPresentationRoundTripsDisplayFlagsDefaultsAndIgnoredErrors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      ExcelSheetPresentationController controller = new ExcelSheetPresentationController();
      ExcelSheetPresentation authored =
          new ExcelSheetPresentation(
              new ExcelSheetDisplay(false, false, false, true, true),
              new ExcelColor("#112233"),
              new ExcelSheetOutlineSummary(false, false),
              new ExcelSheetDefaults(11, 18.5d),
              List.of(
                  new ExcelIgnoredError(
                      "A1:B2",
                      List.of(
                          ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                          ExcelIgnoredErrorType.FORMULA)),
                  new ExcelIgnoredError("D5", List.of(ExcelIgnoredErrorType.TWO_DIGIT_TEXT_YEAR))));

      controller.setPresentation(sheet, authored);

      ExcelSheetPresentationSnapshot snapshot = controller.presentation(sheet);
      assertEquals(authored.display(), snapshot.display());
      assertEquals(new ExcelColorSnapshot("#112233"), snapshot.tabColor());
      assertEquals(authored.outlineSummary(), snapshot.outlineSummary());
      assertEquals(authored.sheetDefaults(), snapshot.sheetDefaults());
      assertEquals(
          List.of(
              new ExcelIgnoredError(
                  "A1:B2",
                  List.of(
                      ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)),
              new ExcelIgnoredError("D5", List.of(ExcelIgnoredErrorType.TWO_DIGIT_TEXT_YEAR))),
          snapshot.ignoredErrors());
      assertTrue(sheet.isRightToLeft());
      assertTrue(sheet.isDisplayFormulas());
      assertFalse(sheet.isDisplayGridlines());
      assertFalse(sheet.isDisplayZeros());
      assertFalse(sheet.isDisplayRowColHeadings());
      assertEquals(11, sheet.getDefaultColumnWidth());
      assertEquals(18.5f, sheet.getDefaultRowHeightInPoints());
      assertEquals(3, sheet.getIgnoredErrors().size());

      controller.clearPresentation(sheet);

      ExcelSheetPresentationSnapshot cleared = controller.presentation(sheet);
      assertEquals(ExcelSheetDisplay.defaults(), cleared.display());
      assertNull(cleared.tabColor());
      assertEquals(ExcelSheetOutlineSummary.defaults(), cleared.outlineSummary());
      assertEquals(ExcelSheetDefaults.defaults(), cleared.sheetDefaults());
      assertEquals(List.of(), cleared.ignoredErrors());
      assertFalse(sheet.isRightToLeft());
      assertFalse(sheet.isDisplayFormulas());
      assertTrue(sheet.isDisplayGridlines());
      assertTrue(sheet.isDisplayZeros());
      assertTrue(sheet.isDisplayRowColHeadings());
      assertNull(sheet.getTabColor());
      assertFalse(sheet.getCTWorksheet().isSetIgnoredErrors());
    }
  }

  @Test
  void presentationFallsBackToDefaultRowHeightWhenWorksheetStoresZero() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      if (!sheet.getCTWorksheet().isSetSheetFormatPr()) {
        sheet.getCTWorksheet().addNewSheetFormatPr();
      }
      sheet.getCTWorksheet().getSheetFormatPr().setDefaultRowHeight(0.0d);

      ExcelSheetPresentationSnapshot snapshot =
          new ExcelSheetPresentationController().presentation(sheet);

      assertEquals(
          ExcelSheetDefaults.defaults().defaultRowHeightPoints(),
          snapshot.sheetDefaults().defaultRowHeightPoints());
    }
  }

  @Test
  void clearPresentationSkipsTabColorUnsetWhenSheetPrHasNoTabColor() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      // Add sheetPr element without a tabColor child so isSetSheetPr() is true but
      // isSetTabColor() is false — exercises the no-op path in clearTabColor.
      if (!sheet.getCTWorksheet().isSetSheetPr()) {
        sheet.getCTWorksheet().addNewSheetPr();
      }

      ExcelSheetPresentationController controller = new ExcelSheetPresentationController();
      controller.clearPresentation(sheet);

      assertNull(sheet.getTabColor());
    }
  }

  @Test
  void setPresentationRejectsSheetDefaultsOutsideExcelLimits() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      ExcelSheetPresentationController controller = new ExcelSheetPresentationController();

      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPresentation(
                  sheet,
                  new ExcelSheetPresentation(
                      ExcelSheetDisplay.defaults(),
                      null,
                      ExcelSheetOutlineSummary.defaults(),
                      new ExcelSheetDefaults(
                          0, ExcelSheetDefaults.defaults().defaultRowHeightPoints()),
                      List.of())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPresentation(
                  sheet,
                  new ExcelSheetPresentation(
                      ExcelSheetDisplay.defaults(),
                      null,
                      ExcelSheetOutlineSummary.defaults(),
                      new ExcelSheetDefaults(
                          ExcelSheetLayoutLimits.MAX_DEFAULT_COLUMN_WIDTH + 1,
                          ExcelSheetDefaults.defaults().defaultRowHeightPoints()),
                      List.of())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPresentation(
                  sheet,
                  new ExcelSheetPresentation(
                      ExcelSheetDisplay.defaults(),
                      null,
                      ExcelSheetOutlineSummary.defaults(),
                      new ExcelSheetDefaults(
                          ExcelSheetDefaults.defaults().defaultColumnWidth(),
                          Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)),
                      List.of())));
    }
  }
}
