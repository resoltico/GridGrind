package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.List;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Direct tests for authoritative sheet-presentation read and write behavior. */
class ExcelSheetPresentationControllerTest {
  @Test
  void normalizesIgnoredErrorCoordinatesOnlyAtPoiBoundary() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Budget");
      var range = new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:A1");
      var authored =
          new SheetPresentationSpec(
              SheetDisplaySpec.defaults(),
              Optional.empty(),
              SheetOutlineSummarySpec.defaults(),
              SheetDefaultsSpec.defaults(),
              List.of(new IgnoredErrorSpec(range, List.of(ExcelIgnoredErrorType.FORMULA))));
      var controller = new ExcelSheetPresentationController();
      controller.setPresentation(sheet, authored);
      assertEquals("B2:A1", range.value());
      assertEquals(
          "A1:B2", controller.presentation(sheet).ignoredErrors().getFirst().range().value());
    }
  }

  @Test
  void setAndClearPresentationRoundTripsDisplayFlagsDefaultsAndIgnoredErrors() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      ExcelSheetPresentationController controller = new ExcelSheetPresentationController();
      SheetPresentationSpec authored =
          new SheetPresentationSpec(
              new SheetDisplaySpec(false, false, false, true, true),
              Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
              new SheetOutlineSummarySpec(false, false),
              new SheetDefaultsSpec(11, 18.5d),
              List.of(
                  new IgnoredErrorSpec(
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B2"),
                      List.of(
                          ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT,
                          ExcelIgnoredErrorType.FORMULA)),
                  new IgnoredErrorSpec(
                      new dev.erst.gridgrind.core.vocabulary.RangeRef("D5"),
                      List.of(ExcelIgnoredErrorType.TWO_DIGIT_TEXT_YEAR))));

      controller.setPresentation(sheet, authored);

      SheetPresentationReport snapshot = controller.presentation(sheet);
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetDisplayReport(
              false, false, false, true, true),
          snapshot.display());
      assertEquals(Optional.of(CellColorReport.rgb(new RgbHex("#112233"))), snapshot.tabColor());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport(false, false),
          snapshot.outlineSummary());
      assertEquals(
          new dev.erst.gridgrind.core.model.layout.SheetDefaultsReport(11, 18.5d),
          snapshot.sheetDefaults());
      assertEquals(
          List.of(
              new dev.erst.gridgrind.core.model.validation.IgnoredErrorReport(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B2"),
                  List.of(
                      ExcelIgnoredErrorType.FORMULA, ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT)),
              new dev.erst.gridgrind.core.model.validation.IgnoredErrorReport(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("D5"),
                  List.of(ExcelIgnoredErrorType.TWO_DIGIT_TEXT_YEAR))),
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

      SheetPresentationReport cleared = controller.presentation(sheet);
      assertEquals(
          dev.erst.gridgrind.core.model.layout.SheetDisplayReport.defaults(), cleared.display());
      assertEquals(Optional.empty(), cleared.tabColor());
      assertEquals(
          dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport.defaults(),
          cleared.outlineSummary());
      assertEquals(
          dev.erst.gridgrind.core.model.layout.SheetDefaultsReport.defaults(),
          cleared.sheetDefaults());
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

      SheetPresentationReport snapshot = new ExcelSheetPresentationController().presentation(sheet);

      assertEquals(
          SheetDefaultsSpec.defaults().defaultRowHeightPoints(),
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
                  new SheetPresentationSpec(
                      SheetDisplaySpec.defaults(),
                      Optional.empty(),
                      SheetOutlineSummarySpec.defaults(),
                      new SheetDefaultsSpec(
                          0, SheetDefaultsSpec.defaults().defaultRowHeightPoints()),
                      List.of())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPresentation(
                  sheet,
                  new SheetPresentationSpec(
                      SheetDisplaySpec.defaults(),
                      Optional.empty(),
                      SheetOutlineSummarySpec.defaults(),
                      new SheetDefaultsSpec(
                          ExcelSheetLayoutLimits.MAX_DEFAULT_COLUMN_WIDTH + 1,
                          SheetDefaultsSpec.defaults().defaultRowHeightPoints()),
                      List.of())));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              controller.setPresentation(
                  sheet,
                  new SheetPresentationSpec(
                      SheetDisplaySpec.defaults(),
                      Optional.empty(),
                      SheetOutlineSummarySpec.defaults(),
                      new SheetDefaultsSpec(
                          SheetDefaultsSpec.defaults().defaultColumnWidth(),
                          Math.nextUp(ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS)),
                      List.of())));
    }
  }
}
