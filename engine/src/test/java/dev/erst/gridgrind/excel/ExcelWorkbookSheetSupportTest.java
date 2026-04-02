package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for workbook-level sheet-state helper methods. */
class ExcelWorkbookSheetSupportTest {
  @Test
  void activeSheetNameRejectsWorkbooksWithoutSheets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelWorkbookSheetSupport.activeSheetName(workbook));
      assertEquals("workbook must contain at least one sheet", exception.getMessage());
    }
  }

  @Test
  void activeSheetNameFallsBackToFirstSheetWhenWorkbookViewStateIsInvalid() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getCTWorkbook().getBookViews().getWorkbookViewArray(0).setActiveTab(99);

      assertEquals("Alpha", ExcelWorkbookSheetSupport.activeSheetName(workbook));
    }
  }

  @Test
  void activeSheetNameFallsBackToFirstSheetWhenWorkbookViewStateIsNegative() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getCTWorkbook().getBookViews().getWorkbookViewArray(0).setActiveTab(-1);

      assertEquals("Alpha", ExcelWorkbookSheetSupport.activeSheetName(workbook));
    }
  }

  @Test
  void selectedSheetNamesHandleEmptyAndMissingSelectionStates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertEquals(List.of(), ExcelWorkbookSheetSupport.selectedSheetNames(workbook));

      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getSheetAt(0).setSelected(false);
      workbook.getSheetAt(1).setSelected(false);
      workbook.setActiveSheet(1);

      assertEquals(List.of("Beta"), ExcelWorkbookSheetSupport.selectedSheetNames(workbook));
    }
  }

  @Test
  void normalizeWorkbookViewStateRepairsHiddenActiveAndHiddenSelections() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.setSheetVisibility(0, SheetVisibility.HIDDEN);
      workbook.setActiveSheet(0);
      workbook.getSheetAt(0).setSelected(true);
      workbook.getSheetAt(1).setSelected(false);

      ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook);

      assertEquals(1, workbook.getActiveSheetIndex());
      assertFalse(workbook.getSheetAt(0).isSelected());
      assertTrue(workbook.getSheetAt(1).isSelected());
    }
  }

  @Test
  void normalizeWorkbookViewStateRepairsNegativeActiveSheetIndexes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getCTWorkbook().getBookViews().getWorkbookViewArray(0).setActiveTab(-1);
      workbook.getSheetAt(0).setSelected(false);
      workbook.getSheetAt(1).setSelected(true);

      ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook);

      assertEquals(0, workbook.getActiveSheetIndex());
      assertTrue(workbook.getSheetAt(0).isSelected());
    }
  }

  @Test
  void normalizeWorkbookViewStateRepairsOutOfRangeActiveSheetIndexes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getCTWorkbook().getBookViews().getWorkbookViewArray(0).setActiveTab(99);
      workbook.getSheetAt(0).setSelected(false);
      workbook.getSheetAt(1).setSelected(true);

      ExcelWorkbookSheetSupport.normalizeWorkbookViewState(workbook);

      assertEquals(0, workbook.getActiveSheetIndex());
      assertTrue(workbook.getSheetAt(0).isSelected());
    }
  }

  @Test
  void firstVisibleSheetIndexRejectsWorkbooksWithoutVisibleSheets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.setSheetVisibility(0, SheetVisibility.HIDDEN);

      IllegalStateException exception =
          assertThrows(
              IllegalStateException.class,
              () -> ExcelWorkbookSheetSupport.firstVisibleSheetIndex(workbook));
      assertEquals("workbook must contain at least one visible sheet", exception.getMessage());
    }
  }

  @Test
  void requireNotLastVisibleSheetRejectsOnlyVisibleSheetTransitions() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.setSheetVisibility(1, SheetVisibility.HIDDEN);

      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  ExcelWorkbookSheetSupport.requireNotLastVisibleSheet(
                      workbook, 0, "cannot remove last visible"));
      assertEquals("cannot remove last visible", exception.getMessage());
      assertDoesNotThrow(
          () -> ExcelWorkbookSheetSupport.requireNotLastVisibleSheet(workbook, 1, "ignored"));
    }
  }

  @Test
  void selectionHelpersValidateNamesAndPresence() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      List<String> selected =
          ExcelWorkbookSheetSupport.copySelectedSheetNames(List.of("Alpha", "Beta"));
      assertEquals(List.of("Alpha", "Beta"), selected);

      List<String> selectedWithNull = new ArrayList<>();
      selectedWithNull.add(null);
      assertThrows(
          NullPointerException.class,
          () -> ExcelWorkbookSheetSupport.copySelectedSheetNames(selectedWithNull));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelWorkbookSheetSupport.copySelectedSheetNames(List.of()));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelWorkbookSheetSupport.copySelectedSheetNames(List.of("Alpha", " ")));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelWorkbookSheetSupport.copySelectedSheetNames(List.of("Alpha", "Alpha")));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelWorkbookSheetSupport.requiredSelectedSheetIndex(workbook, "Missing"));
    }
  }

  @Test
  void requiredSheetHelpersValidateExistenceAndVisibility() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.setSheetVisibility(1, SheetVisibility.HIDDEN);

      assertEquals(
          "Alpha", ExcelWorkbookSheetSupport.requiredSheet(workbook, "Alpha").getSheetName());
      assertEquals(0, ExcelWorkbookSheetSupport.requiredSheetIndex(workbook, "Alpha"));

      assertThrows(
          SheetNotFoundException.class,
          () -> ExcelWorkbookSheetSupport.requiredSheet(workbook, "Missing"));
      assertThrows(
          SheetNotFoundException.class,
          () -> ExcelWorkbookSheetSupport.requiredSheetIndex(workbook, "Missing"));
      IllegalArgumentException hidden =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelWorkbookSheetSupport.requireVisibleSheet(workbook, 1, "no hidden"));
      assertEquals("no hidden", hidden.getMessage());
    }
  }

  @Test
  void nameAvailabilityAndTargetIndexValidationStayHonest() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");

      assertDoesNotThrow(
          () -> ExcelWorkbookSheetSupport.requireSheetNameAvailable(workbook, "alpha", 0));
      IllegalArgumentException collision =
          assertThrows(
              IllegalArgumentException.class,
              () -> ExcelWorkbookSheetSupport.requireSheetNameAvailable(workbook, "ALPHA", 1));
      assertEquals("Sheet already exists: ALPHA", collision.getMessage());

      assertDoesNotThrow(() -> ExcelWorkbookSheetSupport.requireTargetIndex(workbook, 1));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelWorkbookSheetSupport.requireTargetIndex(workbook, -1));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelWorkbookSheetSupport.requireTargetIndex(workbook, 2));
    }
  }

  @Test
  void requireSheetNameValidatesNullBlankAndLength() {
    assertThrows(
        NullPointerException.class,
        () -> ExcelWorkbookSheetSupport.requireSheetName(null, "sheetName"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelWorkbookSheetSupport.requireSheetName(" ", "sheetName"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelWorkbookSheetSupport.requireSheetName(
                "12345678901234567890123456789012", "sheetName"));
  }
}
