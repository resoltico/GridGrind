package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.junit.jupiter.api.Test;

/** Tests for workbook-level sheet-state mutations and summaries. */
class ExcelSheetStateControllerTest {
  @Test
  void setSelectedSheetsRepairsInvalidActiveSheetIndexes() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook
          .xssfWorkbook()
          .getCTWorkbook()
          .getBookViews()
          .getWorkbookViewArray(0)
          .setActiveTab(-1);

      controller.setSelectedSheets(workbook, List.of("Beta"));

      WorkbookReadResult.WorkbookSummary.WithSheets summary =
          assertInstanceOf(
              WorkbookReadResult.WorkbookSummary.WithSheets.class,
              controller.summarizeWorkbook(workbook));
      assertEquals("Beta", summary.activeSheetName());
      assertEquals(List.of("Beta"), summary.selectedSheetNames());
    }
  }

  @Test
  void setSelectedSheetsRepairsOutOfRangeActiveSheetIndexes() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook
          .xssfWorkbook()
          .getCTWorkbook()
          .getBookViews()
          .getWorkbookViewArray(0)
          .setActiveTab(99);

      controller.setSelectedSheets(workbook, List.of("Beta"));

      WorkbookReadResult.WorkbookSummary.WithSheets summary =
          assertInstanceOf(
              WorkbookReadResult.WorkbookSummary.WithSheets.class,
              controller.summarizeWorkbook(workbook));
      assertEquals("Beta", summary.activeSheetName());
      assertEquals(List.of("Beta"), summary.selectedSheetNames());
    }
  }

  @Test
  void setSheetVisibilityAllowsRevealingHiddenSheets() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      controller.setSheetVisibility(workbook, "Beta", ExcelSheetVisibility.HIDDEN);

      controller.setSheetVisibility(workbook, "Beta", ExcelSheetVisibility.VISIBLE);

      assertEquals(
          ExcelSheetVisibility.VISIBLE, controller.summarizeSheet(workbook, "Beta").visibility());
    }
  }

  @Test
  void deleteSheetRejectsDeletingTheLastVisibleSheet() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      controller.setSheetVisibility(workbook, "Beta", ExcelSheetVisibility.HIDDEN);

      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class, () -> controller.deleteSheet(workbook, "Alpha"));
      assertEquals("cannot delete the last visible sheet 'Alpha'", exception.getMessage());
      assertEquals(List.of("Alpha", "Beta"), workbook.sheetNames());
      assertEquals(
          ExcelSheetVisibility.VISIBLE, controller.summarizeSheet(workbook, "Alpha").visibility());
      assertEquals(
          ExcelSheetVisibility.HIDDEN, controller.summarizeSheet(workbook, "Beta").visibility());
    }
  }

  @Test
  void workbookProtectionReadsLockAndPasswordHashFlags() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.xssfWorkbook().lockStructure();
      workbook.xssfWorkbook().lockWindows();
      workbook.xssfWorkbook().lockRevision();
      workbook.xssfWorkbook().setWorkbookPassword("secret", HashAlgorithm.sha512);
      workbook.xssfWorkbook().setRevisionsPassword("review", HashAlgorithm.sha512);
      var protection =
          workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection()
              ? workbook.xssfWorkbook().getCTWorkbook().getWorkbookProtection()
              : workbook.xssfWorkbook().getCTWorkbook().addNewWorkbookProtection();
      protection.setWorkbookPassword(new byte[] {0x01, 0x02});
      protection.setRevisionsPassword(new byte[] {0x03, 0x04});

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(true, true, true, true, true),
          controller.workbookProtection(workbook));
    }
  }

  @Test
  void workbookProtectionDefaultsToAllFalseWithoutWorkbookProtectionXml() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(false, false, false, false, false),
          controller.workbookProtection(workbook));
    }
  }
}
