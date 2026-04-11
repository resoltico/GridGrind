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

  @Test
  void setAndClearWorkbookProtectionRoundTripsAndRemovesEmptyNodes() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      ExcelWorkbookProtectionSettings protectedSettings =
          new ExcelWorkbookProtectionSettings(true, true, true, "secret", "review");

      controller.setWorkbookProtection(workbook, protectedSettings);

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(true, true, true, true, true),
          controller.workbookProtection(workbook));
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());

      controller.clearWorkbookProtection(workbook);

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(false, false, false, false, false),
          controller.workbookProtection(workbook));
      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());

      controller.setWorkbookProtection(
          workbook, new ExcelWorkbookProtectionSettings(false, false, false, null, null));

      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }
  }

  @Test
  void excelWorkbookDelegatesWorkbookProtectionMutations() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      ExcelWorkbookProtectionSettings settings =
          new ExcelWorkbookProtectionSettings(true, false, true, "secret", null);

      workbook.setWorkbookProtection(settings);

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(true, false, true, true, false),
          workbook.workbookProtection());

      workbook.clearWorkbookProtection();

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(false, false, false, false, false),
          workbook.workbookProtection());
    }
  }

  @Test
  void setWorkbookProtectionClearsStaleLegacyAndModernHashesWhenPasswordsAreOmitted()
      throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      var protection = workbook.xssfWorkbook().getCTWorkbook().addNewWorkbookProtection();
      protection.setWorkbookPassword(new byte[] {0x01, 0x02});
      protection.setWorkbookPasswordCharacterSet("1");
      protection.setWorkbookAlgorithmName("SHA-512");
      protection.setWorkbookHashValue(new byte[] {0x03, 0x04});
      protection.setWorkbookSaltValue(new byte[] {0x05, 0x06});
      protection.setWorkbookSpinCount(100000L);
      protection.setRevisionsPassword(new byte[] {0x07, 0x08});
      protection.setRevisionsPasswordCharacterSet("1");
      protection.setRevisionsAlgorithmName("SHA-512");
      protection.setRevisionsHashValue(new byte[] {0x09, 0x0A});
      protection.setRevisionsSaltValue(new byte[] {0x0B, 0x0C});
      protection.setRevisionsSpinCount(100000L);

      controller.setWorkbookProtection(
          workbook, new ExcelWorkbookProtectionSettings(true, false, true, null, null));

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(true, false, true, false, false),
          controller.workbookProtection(workbook));
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
      var normalizedProtection = workbook.xssfWorkbook().getCTWorkbook().getWorkbookProtection();
      assertFalse(normalizedProtection.isSetWorkbookPassword());
      assertFalse(normalizedProtection.isSetWorkbookHashValue());
      assertFalse(normalizedProtection.isSetWorkbookSaltValue());
      assertFalse(normalizedProtection.isSetWorkbookSpinCount());
      assertFalse(normalizedProtection.isSetWorkbookAlgorithmName());
      assertFalse(normalizedProtection.isSetRevisionsPassword());
      assertFalse(normalizedProtection.isSetRevisionsHashValue());
      assertFalse(normalizedProtection.isSetRevisionsSaltValue());
      assertFalse(normalizedProtection.isSetRevisionsSpinCount());
      assertFalse(normalizedProtection.isSetRevisionsAlgorithmName());
    }
  }

  @Test
  void clearWorkbookProtectionIsIdempotentOnFreshWorkbooks() throws IOException {
    ExcelSheetStateController controller = new ExcelSheetStateController();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");

      controller.clearWorkbookProtection(workbook);

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(false, false, false, false, false),
          controller.workbookProtection(workbook));
      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }
  }
}
