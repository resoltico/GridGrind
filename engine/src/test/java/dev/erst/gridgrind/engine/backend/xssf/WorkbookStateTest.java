package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.junit.jupiter.api.Test;

/** Tests for workbook-level sheet-state mutations and summaries. */
class WorkbookStateTest {
  @Test
  void setSelectedSheetsRepairsInvalidActiveSheetIndexes() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook
          .xssfWorkbook()
          .getCTWorkbook()
          .getBookViews()
          .getWorkbookViewArray(0)
          .setActiveTab(-1);

      workbook.sheets().setSelectedSheets(List.of("Beta"));

      WorkbookSummary.WithSheets summary =
          assertInstanceOf(WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals(
          new dev.erst.gridgrind.core.vocabulary.SheetLabel("Beta"), summary.activeSheetName());
      assertEquals(
          List.of(new dev.erst.gridgrind.core.vocabulary.SheetLabel("Beta")),
          summary.selectedSheetNames());
    }
  }

  @Test
  void setSelectedSheetsRepairsOutOfRangeActiveSheetIndexes() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook
          .xssfWorkbook()
          .getCTWorkbook()
          .getBookViews()
          .getWorkbookViewArray(0)
          .setActiveTab(99);

      workbook.sheets().setSelectedSheets(List.of("Beta"));

      WorkbookSummary.WithSheets summary =
          assertInstanceOf(WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals(
          new dev.erst.gridgrind.core.vocabulary.SheetLabel("Beta"), summary.activeSheetName());
      assertEquals(
          List.of(new dev.erst.gridgrind.core.vocabulary.SheetLabel("Beta")),
          summary.selectedSheetNames());
    }
  }

  @Test
  void setSheetVisibilityAllowsRevealingHiddenSheets() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);

      workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.VISIBLE);

      assertEquals(
          ExcelSheetVisibility.VISIBLE, workbook.sheets().summarizeSheet("Beta").visibility());
    }
  }

  @Test
  void deleteSheetRejectsDeletingTheLastVisibleSheet() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.sheets().setSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN);

      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class, () -> workbook.sheets().deleteSheet("Alpha"));
      assertEquals("cannot delete the last visible sheet 'Alpha'", exception.getMessage());
      assertEquals(List.of("Alpha", "Beta"), workbook.sheets().sheetNames());
      assertEquals(
          ExcelSheetVisibility.VISIBLE, workbook.sheets().summarizeSheet("Alpha").visibility());
      assertEquals(
          ExcelSheetVisibility.HIDDEN, workbook.sheets().summarizeSheet("Beta").visibility());
    }
  }

  @Test
  void workbookProtectionReadsLockAndPasswordHashFlags() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
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
          new WorkbookProtectionReport(true, true, true, true, true),
          workbook.protection().snapshotWorkbook());
    }
  }

  @Test
  void workbookProtectionDefaultsToAllFalseWithoutWorkbookProtectionXml() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, false),
          workbook.protection().snapshotWorkbook());
    }
  }

  @Test
  void setAndClearWorkbookProtectionRoundTripsAndRemovesEmptyNodes() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      ResolvedWorkbookProtection protectedSettings =
          new ResolvedWorkbookProtection(
              true, true, true, Optional.of("secret"), Optional.of("review"));

      workbook.protection().setWorkbookProtection(protectedSettings);

      assertEquals(
          new WorkbookProtectionReport(true, true, true, true, true),
          workbook.protection().snapshotWorkbook());
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());

      workbook.protection().clearWorkbookProtection();

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, false),
          workbook.protection().snapshotWorkbook());
      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());

      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  false, false, false, Optional.empty(), Optional.empty()));

      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }
  }

  @Test
  void excelWorkbookDelegatesWorkbookProtectionMutations() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      ResolvedWorkbookProtection settings =
          new ResolvedWorkbookProtection(
              true, false, true, Optional.of("secret"), Optional.empty());

      workbook.protection().setWorkbookProtection(settings);

      assertEquals(
          new WorkbookProtectionReport(true, false, true, true, false),
          workbook.protection().snapshotWorkbook());

      workbook.protection().clearWorkbookProtection();

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, false),
          workbook.protection().snapshotWorkbook());
    }
  }

  @Test
  void setWorkbookProtectionClearsStaleLegacyAndModernHashesWhenPasswordsAreOmitted()
      throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
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

      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  true, false, true, Optional.empty(), Optional.empty()));

      assertEquals(
          new WorkbookProtectionReport(true, false, true, false, false),
          workbook.protection().snapshotWorkbook());
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

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");

      workbook.protection().clearWorkbookProtection();

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, false),
          workbook.protection().snapshotWorkbook());
      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }
  }

  @Test
  void workbookProtectionHashSubconditionsCoverEachIndividualField() throws IOException {

    // normalizeWorkbookProtectionNode: hasLocks — windowsLocked is the first true.
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  false, true, false, Optional.empty(), Optional.empty()));
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }

    // normalizeWorkbookProtectionNode: hasLocks — revisionsLocked is the first true.
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  false, false, true, Optional.empty(), Optional.empty()));
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }

    // workbookPasswordHashPresent: each sub-condition as the first true.
    // Also covers: hasPasswords — workbookPasswordHashPresent=false,
    // revisionsPasswordHashPresent=true.
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      var protection = workbook.xssfWorkbook().getCTWorkbook().addNewWorkbookProtection();

      protection.setWorkbookHashValue(new byte[] {0x01});
      assertTrue(workbook.protection().snapshotWorkbook().workbookPasswordHashPresent());
      protection.unsetWorkbookHashValue();

      protection.setWorkbookSaltValue(new byte[] {0x02});
      assertTrue(workbook.protection().snapshotWorkbook().workbookPasswordHashPresent());
      protection.unsetWorkbookSaltValue();

      protection.setWorkbookSpinCount(100_000L);
      assertTrue(workbook.protection().snapshotWorkbook().workbookPasswordHashPresent());
      protection.unsetWorkbookSpinCount();

      protection.setWorkbookAlgorithmName("SHA-512");
      assertTrue(workbook.protection().snapshotWorkbook().workbookPasswordHashPresent());
      protection.unsetWorkbookAlgorithmName();

      // revisionsPasswordHashPresent sub-conditions.
      protection.setRevisionsHashValue(new byte[] {0x03});
      assertTrue(workbook.protection().snapshotWorkbook().revisionsPasswordHashPresent());
      protection.unsetRevisionsHashValue();

      protection.setRevisionsSaltValue(new byte[] {0x04});
      assertTrue(workbook.protection().snapshotWorkbook().revisionsPasswordHashPresent());
      protection.unsetRevisionsSaltValue();

      protection.setRevisionsSpinCount(100_000L);
      assertTrue(workbook.protection().snapshotWorkbook().revisionsPasswordHashPresent());
      protection.unsetRevisionsSpinCount();

      protection.setRevisionsAlgorithmName("SHA-512");
      assertTrue(workbook.protection().snapshotWorkbook().revisionsPasswordHashPresent());
      protection.unsetRevisionsAlgorithmName();

      // hasPasswords: workbookPasswordHashPresent=false, revisionsPasswordHashPresent=true.
      protection.setRevisionsPassword(new byte[] {0x05});
      assertFalse(workbook.protection().snapshotWorkbook().workbookPasswordHashPresent());
      assertTrue(workbook.protection().snapshotWorkbook().revisionsPasswordHashPresent());
    }
  }

  @Test
  void setWorkbookProtectionReusesExistingNodeAndRemovesItWhenNothingRemains() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.xssfWorkbook().getCTWorkbook().addNewWorkbookProtection();

      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  false, false, false, Optional.empty(), Optional.empty()));

      assertFalse(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }
  }

  @Test
  void setWorkbookProtectionRetainsNodeWhenOnlyPasswordHashesRemain() throws IOException {

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");

      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  false, false, false, Optional.of("secret"), Optional.empty()));

      assertEquals(
          new WorkbookProtectionReport(false, false, false, true, false),
          workbook.protection().snapshotWorkbook());
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");

      workbook
          .protection()
          .setWorkbookProtection(
              new ResolvedWorkbookProtection(
                  false, false, false, Optional.empty(), Optional.of("review")));

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, true),
          workbook.protection().snapshotWorkbook());
      assertTrue(workbook.xssfWorkbook().getCTWorkbook().isSetWorkbookProtection());
    }
  }
}
