package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import org.junit.jupiter.api.Test;

/** Tests for the table-header metadata synchronization support seam. */
class ExcelTableHeaderSyncSupportTest {
  private final TableAdapter tableController = new TableAdapter();

  @Test
  void syncAffectedHeaders_updatesPersistedMetadataWhenHeaderRangeIsTouched() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ResolvedTable(
              new TableName("Queue"),
              new SheetName("Ops"),
              new RangeRef("A1:B3"),
              false,
              new TableStyleSpec.None()));

      var poiSheet = workbook.sheet("Ops").xssfSheet();
      var table = poiSheet.getTables().getFirst();
      poiSheet.getRow(0).getCell(0).setCellValue("Desk");

      TableAdapter.syncAffectedHeaders(poiSheet, new CellBounds(0, 0, 0, 0));

      assertEquals("Desk", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  @Test
  void syncAffectedHeaders_ignoresRangesOutsideHeaderRows() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ResolvedTable(
              new TableName("Queue"),
              new SheetName("Ops"),
              new RangeRef("A1:B3"),
              false,
              new TableStyleSpec.None()));

      var poiSheet = workbook.sheet("Ops").xssfSheet();
      var table = poiSheet.getTables().getFirst();
      poiSheet.getRow(1).getCell(0).setCellValue("Grace");

      TableAdapter.syncAffectedHeaders(poiSheet, new CellBounds(1, 1, 0, 0));

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  @Test
  void syncAffectedHeaders_ignoresHeaderlessAndInvalidTables() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ResolvedTable(
              new TableName("Queue"),
              new SheetName("Ops"),
              new RangeRef("A1:B3"),
              false,
              new TableStyleSpec.None()));

      var poiSheet = workbook.sheet("Ops").xssfSheet();
      var table = poiSheet.getTables().getFirst();
      poiSheet.getRow(0).getCell(0).setCellValue("Desk");
      table.getCTTable().setHeaderRowCount(0);

      TableAdapter.syncAffectedHeaders(poiSheet, new CellBounds(0, 0, 0, 0));

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());

      table.getCTTable().setHeaderRowCount(1);
      table.getCTTable().setRef("");

      TableAdapter.syncAffectedHeaders(poiSheet, new CellBounds(0, 0, 0, 0));

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  @Test
  void syncAllHeaders_ignoresHeaderlessAndInvalidTables() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet sheet = workbook.getOrCreateSheet("Ops");
      populateTableCells(sheet, "Owner", "Task");
      tableController.setTable(
          workbook,
          new ResolvedTable(
              new TableName("Queue"),
              new SheetName("Ops"),
              new RangeRef("A1:B3"),
              false,
              new TableStyleSpec.None()));

      var table = workbook.sheet("Ops").xssfSheet().getTables().getFirst();
      workbook.sheet("Ops").xssfSheet().getRow(0).getCell(0).setCellValue("Desk");
      table.getCTTable().setHeaderRowCount(0);

      TableAdapter.syncAllHeaders(workbook.xssfWorkbook());

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());

      table.getCTTable().setHeaderRowCount(1);
      table.getCTTable().setRef("");

      TableAdapter.syncAllHeaders(workbook.xssfWorkbook());

      assertEquals("Owner", table.getCTTable().getTableColumns().getTableColumnArray(0).getName());
    }
  }

  private static void populateTableCells(XssfSheet sheet, String firstHeader, String secondHeader) {
    sheet.cells().setCell("A1", new CellValue.Text(firstHeader));
    sheet.cells().setCell("B1", new CellValue.Text(secondHeader));
    sheet.cells().setCell("A2", new CellValue.Text("Ada"));
    sheet.cells().setCell("B2", new CellValue.Text("Queue"));
    sheet.cells().setCell("A3", new CellValue.Text("Lin"));
    sheet.cells().setCell("B3", new CellValue.Text("Pack"));
  }
}
