package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for package-private table catalog and snapshot helpers. */
class ExcelTableCatalogSupportTest {
  @Test
  void selectTablesByNamePreservesOrderAndIgnoresMissingNames() {
    List<TableEntryReport> allTables =
        List.of(
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                0,
                List.of("Owner", "Task"),
                new TableStyleReport.None(),
                true),
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("ArchiveQueue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Archive"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                0,
                List.of("Owner", "Task"),
                new TableStyleReport.None(),
                false));

    List<TableEntryReport> selected =
        TableAdapter.selectTablesByName(allTables, List.of("archivequeue", "missing", "QUEUE"));

    assertEquals(
        List.of("ArchiveQueue", "Queue"),
        selected.stream().map(table -> table.name().value()).toList());
    assertEquals(Optional.empty(), TableAdapter.findTableSnapshotByName(allTables, "missing"));
  }

  @Test
  void requiredTableByNameFindsExistingTablesAndRejectsMissingOnes() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Owner");
      sheet.getRow(0).createCell(1).setCellValue("Task");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");
      XSSFTable table = sheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
      table.setName("Queue");

      assertSame(table, TableAdapter.requiredTableByName(sheet, "queue"));
      assertThrows(
          IllegalArgumentException.class, () -> TableAdapter.requiredTableByName(sheet, "Missing"));
    }
  }

  @Test
  void toSnapshotCapturesStyleAndAutofilterMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Owner");
      sheet.getRow(0).createCell(1).setCellValue("Task");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");
      sheet.createRow(2).createCell(0).setCellValue("Lin");
      sheet.getRow(2).createCell(1).setCellValue("Pack");
      XSSFTable table = sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));
      table.setName("Queue");
      TableAdapter.applyAutofilter(table, CellBounds.parse("A1:B3"), false);
      TableAdapter.applyStyle(
          table, new TableStyleSpec.Named("TableStyleMedium2", true, false, true, false));
      table.getCTTable().setComment("Queue comment");
      table.getCTTable().setPublished(true);
      table.getCTTable().setInsertRow(true);
      table.getCTTable().setInsertRowShift(false);
      table.getCTTable().setHeaderRowCellStyle("HeaderStyle");
      table.getCTTable().setDataCellStyle("DataStyle");
      table.getCTTable().setTotalsRowCellStyle("TotalsStyle");
      var amountColumn = table.getCTTable().getTableColumns().getTableColumnArray(1);
      amountColumn.setUniqueName("UniqueAmount");
      amountColumn.setTotalsRowLabel("Total");
      amountColumn.setTotalsRowFunction(
          org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction.SUM);
      amountColumn.addNewCalculatedColumnFormula().setStringValue("[@Task]&\"!\"");

      TableEntryReport snapshot = TableAdapter.toSnapshot("Ops", table);

      assertEquals("Queue", snapshot.name().value());
      assertEquals(List.of("Owner", "Task"), snapshot.structure().columnNames());
      assertTrue(snapshot.behavior().hasAutofilter());
      assertInstanceOf(TableStyleReport.Named.class, snapshot.style());
      assertEquals("Queue comment", snapshot.presentation().comment().orElseThrow());
      assertEquals("HeaderStyle", snapshot.presentation().headerRowCellStyle().orElseThrow());
      assertEquals(
          "UniqueAmount", snapshot.structure().columns().get(1).uniqueName().orElseThrow());
      assertEquals("Total", snapshot.structure().columns().get(1).totalsRowLabel().orElseThrow());
      assertEquals("sum", snapshot.structure().columns().get(1).totalsRowFunction().orElseThrow());
      assertEquals(
          "[@Task]&\"!\"",
          snapshot.structure().columns().get(1).calculatedColumnFormula().orElseThrow());
    }
  }

  @Test
  void toSnapshotNormalizesNullAndBlankPresentationMetadataToEmptyOptionals() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(0).createCell(0).setCellValue("Owner");
      sheet.getRow(0).createCell(1).setCellValue("Task");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");
      XSSFTable table = sheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
      table.setName("Queue");
      table.getCTTable().setHeaderRowCellStyle(" ");
      table.getCTTable().setDataCellStyle("DataStyle");
      table.getCTTable().setTotalsRowCellStyle("");

      TableEntryReport snapshot = TableAdapter.toSnapshot("Ops", table);

      assertTrue(snapshot.presentation().comment().isEmpty());
      assertTrue(snapshot.presentation().headerRowCellStyle().isEmpty());
      assertEquals("DataStyle", snapshot.presentation().dataCellStyle().orElseThrow());
      assertTrue(snapshot.presentation().totalsRowCellStyle().isEmpty());
    }
  }
}
