package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for workbook-level sheet-state helper methods. */
class SheetCollectionInvariantsTest {
  @Test
  void activeSheetNameRejectsWorkbooksWithoutSheets() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () -> SheetCollectionAdapter.activeSheetName(workbook));
      assertEquals("workbook must contain at least one sheet", exception.getMessage());
    }
  }

  @Test
  void activeSheetNameFallsBackToFirstSheetWhenWorkbookViewStateIsInvalid() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getCTWorkbook().getBookViews().getWorkbookViewArray(0).setActiveTab(99);

      assertEquals("Alpha", SheetCollectionAdapter.activeSheetName(workbook));
    }
  }

  @Test
  void activeSheetNameFallsBackToFirstSheetWhenWorkbookViewStateIsNegative() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getCTWorkbook().getBookViews().getWorkbookViewArray(0).setActiveTab(-1);

      assertEquals("Alpha", SheetCollectionAdapter.activeSheetName(workbook));
    }
  }

  @Test
  void selectedSheetNamesHandleEmptyAndMissingSelectionStates() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      assertEquals(List.of(), SheetCollectionAdapter.selectedSheetNames(workbook));

      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.getSheetAt(0).setSelected(false);
      workbook.getSheetAt(1).setSelected(false);
      workbook.setActiveSheet(1);

      assertEquals(List.of("Beta"), SheetCollectionAdapter.selectedSheetNames(workbook));
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

      SheetCollectionAdapter.normalizeWorkbookViewState(workbook);

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

      SheetCollectionAdapter.normalizeWorkbookViewState(workbook);

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

      SheetCollectionAdapter.normalizeWorkbookViewState(workbook);

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
              () -> SheetCollectionAdapter.firstVisibleSheetIndex(workbook));
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
                  SheetCollectionAdapter.requireNotLastVisibleSheet(
                      workbook, 0, "cannot remove last visible"));
      assertEquals("cannot remove last visible", exception.getMessage());
      assertDoesNotThrow(
          () -> SheetCollectionAdapter.requireNotLastVisibleSheet(workbook, 1, "ignored"));
    }
  }

  @Test
  void selectionHelpersValidateNamesAndPresence() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      List<String> selected =
          SheetCollectionAdapter.copySelectedSheetNames(List.of("Alpha", "Beta"));
      assertEquals(List.of("Alpha", "Beta"), selected);

      List<String> selectedWithNull = new ArrayList<>();
      selectedWithNull.add(null);
      assertThrows(
          NullPointerException.class,
          () -> SheetCollectionAdapter.copySelectedSheetNames(selectedWithNull));
      assertThrows(
          IllegalArgumentException.class,
          () -> SheetCollectionAdapter.copySelectedSheetNames(List.of()));
      assertThrows(
          IllegalArgumentException.class,
          () -> SheetCollectionAdapter.copySelectedSheetNames(List.of("Alpha", " ")));
      assertThrows(
          IllegalArgumentException.class,
          () -> SheetCollectionAdapter.copySelectedSheetNames(List.of("Alpha", "Alpha")));
      assertThrows(
          IllegalArgumentException.class,
          () -> SheetCollectionAdapter.requiredSelectedSheetIndex(workbook, "Missing"));
    }
  }

  @Test
  void requiredSheetHelpersValidateExistenceAndVisibility() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");
      workbook.setSheetVisibility(1, SheetVisibility.HIDDEN);

      assertEquals("Alpha", SheetCollectionAdapter.requiredSheet(workbook, "Alpha").getSheetName());
      assertEquals(0, SheetCollectionAdapter.requiredSheetIndex(workbook, "Alpha"));

      assertThrows(
          SheetNotFoundException.class,
          () -> SheetCollectionAdapter.requiredSheet(workbook, "Missing"));
      assertThrows(
          SheetNotFoundException.class,
          () -> SheetCollectionAdapter.requiredSheetIndex(workbook, "Missing"));
      IllegalArgumentException hidden =
          assertThrows(
              IllegalArgumentException.class,
              () -> SheetCollectionAdapter.requireVisibleSheet(workbook, 1, "no hidden"));
      assertEquals("no hidden", hidden.getMessage());
    }
  }

  @Test
  void nameAvailabilityAndTargetIndexValidationStayHonest() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Alpha");
      workbook.createSheet("Beta");

      assertDoesNotThrow(
          () -> SheetCollectionAdapter.requireSheetNameAvailable(workbook, "alpha", 0));
      IllegalArgumentException collision =
          assertThrows(
              IllegalArgumentException.class,
              () -> SheetCollectionAdapter.requireSheetNameAvailable(workbook, "ALPHA", 1));
      assertEquals("Sheet already exists: ALPHA", collision.getMessage());

      assertDoesNotThrow(() -> SheetCollectionAdapter.requireTargetIndex(workbook, 1));
      assertThrows(
          IllegalArgumentException.class,
          () -> SheetCollectionAdapter.requireTargetIndex(workbook, -1));
      assertThrows(
          IllegalArgumentException.class,
          () -> SheetCollectionAdapter.requireTargetIndex(workbook, 2));
    }
  }

  @Test
  void requireSheetNameValidatesNullBlankAndLength() {
    assertThrows(
        NullPointerException.class,
        () -> SheetCollectionAdapter.requireSheetName(null, "sheetName"));
    assertThrows(
        IllegalArgumentException.class,
        () -> SheetCollectionAdapter.requireSheetName(" ", "sheetName"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            SheetCollectionAdapter.requireSheetName(
                "12345678901234567890123456789012", "sheetName"));
  }
}
