package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;

/** Outline grouping, collapse, and sparse-sheet layout coverage. */
class ExcelOutlineGroupingCoverageTest extends ExcelRowColumnStructureTestSupport {
  @Test
  void layoutReportsHiddenOutlineAndCollapsedFacts() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), true);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), true);
      controller.setRowVisibility(sheet, new ExcelRowSpan(5, 5), true);
      controller.setColumnVisibility(sheet, new ExcelColumnSpan(5, 5), true);

      List<WorkbookSheetResult.RowLayout> rows = controller.rowLayouts(sheet);
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);

      assertEquals(6, rows.size());
      assertTrue(rows.get(1).hidden());
      assertEquals(1, rows.get(1).outlineLevel());
      assertTrue(rows.get(4).collapsed());
      assertTrue(rows.get(5).hidden());

      assertEquals(6, columns.size());
      assertTrue(columns.get(1).hidden());
      assertEquals(1, columns.get(1).outlineLevel());
      assertTrue(columns.get(4).collapsed());
      assertTrue(columns.get(5).hidden());
    }
  }

  @Test
  void collapsedRowGroupsPersistControlRowsAtSparseSheetTail() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-row-group-tail-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.groupRows(sheet, new ExcelRowSpan(0, 1), true);

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookSheetResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(3, layout.rows().size());
    assertTrue(layout.rows().get(0).hidden());
    assertEquals(1, layout.rows().get(0).outlineLevel());
    assertTrue(layout.rows().get(1).hidden());
    assertEquals(1, layout.rows().get(1).outlineLevel());
    assertFalse(layout.rows().get(2).hidden());
    assertTrue(layout.rows().get(2).collapsed());
  }

  @Test
  void redundantNoOpColumnUngroupsDoNotMaterializeGhostColumnMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));

      assertTrue(controller.columnLayouts(sheet).isEmpty());
      assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());
      assertEquals(0, sheet.getCTWorksheet().getColsArray(0).sizeOfColArray());
    }
  }

  @Test
  void partialColumnUngroupPreservesRemainingCollapsedBandAcrossRoundTrip() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-column-ungroup-collapse-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(0, 3));
      controller.groupColumns(sheet, new ExcelColumnSpan(0, 3), true);
      controller.ungroupColumns(sheet, new ExcelColumnSpan(1, 1));

      List<WorkbookSheetResult.ColumnLayout> inMemoryColumns = controller.columnLayouts(sheet);
      assertEquals(5, inMemoryColumns.size(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(0).hidden(), "in-memory columns=" + inMemoryColumns);
      assertFalse(inMemoryColumns.get(1).collapsed(), "in-memory columns=" + inMemoryColumns);
      assertEquals(
          0, inMemoryColumns.get(1).outlineLevel(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(2).hidden(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(3).hidden(), "in-memory columns=" + inMemoryColumns);
      assertTrue(inMemoryColumns.get(4).collapsed(), "in-memory columns=" + inMemoryColumns);

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookSheetResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(5, layout.columns().size(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(0).hidden(), "reopened columns=" + layout.columns());
    assertFalse(layout.columns().get(1).collapsed(), "reopened columns=" + layout.columns());
    assertEquals(0, layout.columns().get(1).outlineLevel(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(2).hidden(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(3).hidden(), "reopened columns=" + layout.columns());
    assertTrue(layout.columns().get(4).collapsed(), "reopened columns=" + layout.columns());
  }

  @Test
  void ungroupedSparseRowsNormalizeOutlineLevelToZeroAfterReopen() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-row-ungroup-tail-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.ungroupRows(sheet, new ExcelRowSpan(1, 3));

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookSheetResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(4, layout.rows().size());
    assertEquals(0, layout.rows().get(1).outlineLevel());
    assertEquals(0, layout.rows().get(2).outlineLevel());
    assertEquals(0, layout.rows().get(3).outlineLevel());
  }

  @Test
  void groupRowsAndColumnsWithoutCollapseKeepsBandsExpanded() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), false);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), false);

      List<WorkbookSheetResult.RowLayout> rows = controller.rowLayouts(sheet);
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);

      assertEquals(4, rows.size());
      assertFalse(rows.get(1).hidden());
      assertEquals(1, rows.get(1).outlineLevel());
      assertFalse(rows.get(3).collapsed());

      assertEquals(4, columns.size());
      assertFalse(columns.get(1).hidden());
      assertEquals(1, columns.get(1).outlineLevel());
      assertFalse(columns.get(3).collapsed());
    }
  }

  @Test
  void regroupingCollapsedRowsWithExpandedIntentReopensTheExistingBand() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");
      setString(sheet, "A2", "North");
      setString(sheet, "A3", "South");
      setString(sheet, "A4", "West");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), true);
      assertTrue(controller.rowLayouts(sheet).get(1).hidden());

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), false);

      List<WorkbookSheetResult.RowLayout> rows = controller.rowLayouts(sheet);

      assertEquals(5, rows.size());
      assertFalse(rows.get(1).hidden());
      assertFalse(rows.get(4).collapsed());
    }
  }

  @Test
  void groupingExpandedRowsIgnoresAnExistingUncollapsedControlRow() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");
      setString(sheet, "A2", "North");
      setString(sheet, "A3", "South");
      setString(sheet, "A4", "West");
      setString(sheet, "A5", "Tail");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), false);

      List<WorkbookSheetResult.RowLayout> rows = controller.rowLayouts(sheet);

      assertEquals(5, rows.size());
      assertFalse(rows.get(1).hidden());
      assertFalse(rows.get(4).collapsed());
    }
  }

  @Test
  void expandedRowGroupsPreserveManualHiddenRowsWhenGroupingAndUngrouping() throws Exception {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-row-group-hidden-");
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");
      setString(sheet, "A2", "Visible");
      setString(sheet, "A3", "Hidden");
      setString(sheet, "A4", "Visible");

      controller.setRowVisibility(sheet, new ExcelRowSpan(2, 2), true);
      controller.groupRows(sheet, new ExcelRowSpan(1, 3), false);

      List<WorkbookSheetResult.RowLayout> groupedRows = controller.rowLayouts(sheet);
      assertEquals(4, groupedRows.size());
      assertFalse(groupedRows.get(1).hidden());
      assertTrue(groupedRows.get(2).hidden());
      assertEquals(1, groupedRows.get(2).outlineLevel());
      assertFalse(groupedRows.get(3).collapsed());

      controller.ungroupRows(sheet, new ExcelRowSpan(1, 3));

      List<WorkbookSheetResult.RowLayout> ungroupedRows = controller.rowLayouts(sheet);
      assertTrue(ungroupedRows.get(2).hidden());
      assertEquals(0, ungroupedRows.get(2).outlineLevel());

      try (var output = Files.newOutputStream(workbookPath)) {
        workbook.write(output);
      }
    }

    WorkbookSheetResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Budget");

    assertEquals(4, layout.rows().size());
    assertFalse(layout.rows().get(1).hidden());
    assertTrue(layout.rows().get(2).hidden());
    assertEquals(0, layout.rows().get(2).outlineLevel());
    assertFalse(layout.rows().get(2).collapsed());
  }

  @Test
  void ungroupRowsAndColumnsExpandCollapsedBandsBeforeRemovingOutline() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Header");

      controller.groupRows(sheet, new ExcelRowSpan(1, 3), true);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 3), true);

      controller.ungroupRows(sheet, new ExcelRowSpan(1, 3));
      controller.ungroupColumns(sheet, new ExcelColumnSpan(1, 3));

      List<WorkbookSheetResult.RowLayout> rows = controller.rowLayouts(sheet);
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);

      assertFalse(rows.get(1).hidden());
      assertEquals(0, rows.get(1).outlineLevel());
      assertFalse(columns.get(1).hidden());
      assertEquals(0, columns.get(1).outlineLevel());
    }
  }

  @Test
  void groupingAtWorkbookLimitsSkipsControlArtifacts() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet rowLimitSheet = workbook.createSheet("RowLimit");
      rowLimitSheet.createRow(ExcelRowSpan.MAX_ROW_INDEX).createCell(0).setCellValue("Edge");

      assertDoesNotThrow(
          () ->
              controller.groupRows(
                  rowLimitSheet,
                  new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX, ExcelRowSpan.MAX_ROW_INDEX),
                  true));
      assertDoesNotThrow(
          () ->
              controller.ungroupRows(
                  rowLimitSheet,
                  new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX, ExcelRowSpan.MAX_ROW_INDEX)));

      XSSFSheet columnLimitSheet = workbook.createSheet("ColumnLimit");
      setString(columnLimitSheet, "XFD1", "Edge");

      assertDoesNotThrow(
          () ->
              controller.groupColumns(
                  columnLimitSheet,
                  new ExcelColumnSpan(
                      ExcelColumnSpan.MAX_COLUMN_INDEX, ExcelColumnSpan.MAX_COLUMN_INDEX),
                  false));
      assertDoesNotThrow(
          () ->
              controller.ungroupColumns(
                  columnLimitSheet,
                  new ExcelColumnSpan(
                      ExcelColumnSpan.MAX_COLUMN_INDEX, ExcelColumnSpan.MAX_COLUMN_INDEX)));
    }
  }

  @Test
  void columnStructuralEditsPreserveOnePoiColumnContainerWithoutExplicitMetadata()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet insertSheet = workbook.createSheet("Insert");
      setString(insertSheet, "A1", "Left");
      setString(insertSheet, "B1", "Right");

      controller.insertColumns(insertSheet, 1, 1);

      assertEquals(1, insertSheet.getCTWorksheet().sizeOfColsArray());

      XSSFSheet shiftSheet = workbook.createSheet("Shift");
      setString(shiftSheet, "A1", "Left");
      setString(shiftSheet, "B1", "Middle");
      setString(shiftSheet, "C1", "Right");

      controller.shiftColumns(shiftSheet, new ExcelColumnSpan(0, 1), 1);

      assertEquals(1, shiftSheet.getCTWorksheet().sizeOfColsArray());

      XSSFSheet deleteSheet = workbook.createSheet("Delete");
      setString(deleteSheet, "A1", "Left");
      setString(deleteSheet, "B1", "Drop");
      setString(deleteSheet, "C1", "Keep");

      controller.deleteColumns(deleteSheet, new ExcelColumnSpan(1, 1));

      assertEquals(1, deleteSheet.getCTWorksheet().sizeOfColsArray());
    }
  }

  @Test
  void groupColumnsNormalizesMissingPoiColumnContainerBeforeGrouping() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "B1", "Middle");
      setString(sheet, "C1", "Right");
      sheet.getCTWorksheet().setColsArray(new CTCols[0]);

      controller.groupColumns(sheet, new ExcelColumnSpan(1, 2), false);

      assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(1, columns.get(1).outlineLevel());
      assertEquals(1, columns.get(2).outlineLevel());
      assertFalse(columns.get(1).hidden());
      assertFalse(columns.get(2).hidden());
    }
  }

  @Test
  void groupColumnsStillWorksAfterColumnStructuralEditsWithoutExplicitMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "B1", "Middle");
      setString(sheet, "C1", "Right");

      controller.shiftColumns(sheet, new ExcelColumnSpan(0, 1), 1);
      controller.insertColumns(sheet, 0, 1);
      controller.groupColumns(sheet, new ExcelColumnSpan(1, 2), false);

      assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());
      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(1, columns.get(1).outlineLevel());
      assertEquals(1, columns.get(2).outlineLevel());
      assertFalse(columns.get(1).hidden());
      assertFalse(columns.get(2).hidden());
    }
  }

  @Test
  void groupColumnsClearsCollapsedStateOnlyOnTheMatchingControlColumn() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "C1", "Control");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(2, 4096);

      controller.groupColumns(sheet, new ExcelColumnSpan(1, 1), false);

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(8.0d, columns.get(0).widthCharacters());
      assertFalse(columns.get(2).collapsed());
    }
  }

  @Test
  void expandedOuterColumnGroupingDoesNotExpandCollapsedDescendantsOrCrash() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");

      controller.groupColumns(sheet, new ExcelColumnSpan(2, 3), true);
      assertDoesNotThrow(() -> controller.groupColumns(sheet, new ExcelColumnSpan(0, 2), false));

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertEquals(5, columns.size(), "columns=" + columns);
      assertFalse(columns.get(0).hidden(), "columns=" + columns);
      assertFalse(columns.get(1).hidden(), "columns=" + columns);
      assertTrue(columns.get(2).hidden(), "columns=" + columns);
      assertEquals(2, columns.get(2).outlineLevel(), "columns=" + columns);
      assertTrue(columns.get(3).hidden(), "columns=" + columns);
      assertEquals(1, columns.get(3).outlineLevel(), "columns=" + columns);
      assertTrue(columns.get(4).collapsed(), "columns=" + columns);
    }
  }

  @Test
  void shortColumnGroupingSequencesAvoidUnexpectedPoiRuntimeFailures() throws Exception {
    List<ColumnGroupingMutation> mutations = columnGroupingMutations();

    for (ColumnGroupingMutation first : mutations) {
      for (ColumnGroupingMutation second : mutations) {
        assertUnexpectedRuntimeFailureAbsent(first, second);
      }
    }
  }

  @Test
  void setColumnCollapsedSkipsEarlierRangesBeforeMatchingTheTargetRange() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "A1", "Left");
      setString(sheet, "D1", "Right");
      sheet.setColumnWidth(0, 2048);
      sheet.setColumnWidth(3, 4096);

      ExcelRowColumnStructureController.setColumnCollapsed(sheet, 3, true);

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertFalse(columns.get(0).collapsed());
      assertTrue(columns.get(3).collapsed());
    }
  }

  @Test
  void setColumnCollapsedIgnoresLaterRangesWhenTheTargetComesFirst() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      setString(sheet, "D1", "Right");
      sheet.setColumnWidth(3, 4096);

      ExcelRowColumnStructureController.setColumnCollapsed(sheet, 0, true);

      List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);
      assertFalse(columns.get(3).collapsed());
    }
  }
}
