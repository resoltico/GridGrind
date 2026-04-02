package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/** Tests for the table-header metadata synchronization support seam. */
class ExcelTableHeaderSyncSupportTest {
  private final ExcelTableController tableController = new ExcelTableController();

  @Test
  void syncAffectedHeaders_updatesPersistedMetadataWhenHeaderRangeIsTouched() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      var poiSheet = workbook.sheet("Ops").xssfSheet();
      var table = poiSheet.getTables().getFirst();
      poiSheet.getRow(0).getCell(0).setCellValue("Desk");

      ExcelTableHeaderSyncSupport.syncAffectedHeaders(poiSheet, new ExcelRange(0, 0, 0, 0));

      assertEquals("Desk", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  @Test
  void syncAffectedHeaders_ignoresRangesOutsideHeaderRows() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      var poiSheet = workbook.sheet("Ops").xssfSheet();
      var table = poiSheet.getTables().getFirst();
      poiSheet.getRow(1).getCell(0).setCellValue("Grace");

      ExcelTableHeaderSyncSupport.syncAffectedHeaders(poiSheet, new ExcelRange(1, 1, 0, 0));

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  @Test
  void syncAffectedHeaders_ignoresHeaderlessAndInvalidTables() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      var poiSheet = workbook.sheet("Ops").xssfSheet();
      var table = poiSheet.getTables().getFirst();
      poiSheet.getRow(0).getCell(0).setCellValue("Desk");
      table.getCTTable().setHeaderRowCount(0);

      ExcelTableHeaderSyncSupport.syncAffectedHeaders(poiSheet, new ExcelRange(0, 0, 0, 0));

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());

      table.getCTTable().setHeaderRowCount(1);
      table.getCTTable().setRef("");

      ExcelTableHeaderSyncSupport.syncAffectedHeaders(poiSheet, new ExcelRange(0, 0, 0, 0));

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  @Test
  void syncAllHeaders_ignoresHeaderlessAndInvalidTables() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ExcelTableDefinition("Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None()));

      var table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      workbook.sheet("Ops").xssfSheet().getRow(0).getCell(0).setCellValue("Desk");
      table.getCTTable().setHeaderRowCount(0);

      ExcelTableHeaderSyncSupport.syncAllHeaders(workbook.xssfWorkbook());

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());

      table.getCTTable().setHeaderRowCount(1);
      table.getCTTable().setRef("");

      ExcelTableHeaderSyncSupport.syncAllHeaders(workbook.xssfWorkbook());

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  private static void populateTableCells(
      ExcelSheet sheet, String firstHeader, String secondHeader) {
    sheet.setCell("A1", ExcelCellValue.text(firstHeader));
    sheet.setCell("B1", ExcelCellValue.text(secondHeader));
    sheet.setCell("A2", ExcelCellValue.text("Ada"));
    sheet.setCell("B2", ExcelCellValue.text("Queue"));
    sheet.setCell("A3", ExcelCellValue.text("Lin"));
    sheet.setCell("B3", ExcelCellValue.text("Pack"));
  }
}
