package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
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
    List<ExcelTableSnapshot> allTables =
        List.of(
            new ExcelTableSnapshot(
                "Queue",
                "Ops",
                "A1:B3",
                1,
                0,
                List.of("Owner", "Task"),
                new ExcelTableStyleSnapshot.None(),
                true),
            new ExcelTableSnapshot(
                "ArchiveQueue",
                "Archive",
                "A1:B3",
                1,
                0,
                List.of("Owner", "Task"),
                new ExcelTableStyleSnapshot.None(),
                false));

    List<ExcelTableSnapshot> selected =
        ExcelTableCatalogSupport.selectTablesByName(
            allTables, List.of("archivequeue", "missing", "QUEUE"));

    assertEquals(
        List.of("ArchiveQueue", "Queue"), selected.stream().map(ExcelTableSnapshot::name).toList());
    assertNull(ExcelTableCatalogSupport.findTableSnapshotByName(allTables, "missing"));
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

      assertSame(table, ExcelTableCatalogSupport.requiredTableByName(sheet, "queue"));
      assertThrows(
          IllegalArgumentException.class,
          () -> ExcelTableCatalogSupport.requiredTableByName(sheet, "Missing"));
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
      ExcelTableStructureSupport.applyAutofilter(table, ExcelRange.parse("A1:B3"), false);
      ExcelTableStructureSupport.applyStyle(
          table, new ExcelTableStyle.Named("TableStyleMedium2", true, false, true, false));

      ExcelTableSnapshot snapshot = ExcelTableCatalogSupport.toSnapshot("Ops", table);

      assertEquals("Queue", snapshot.name());
      assertEquals(List.of("Owner", "Task"), snapshot.columnNames());
      assertTrue(snapshot.hasAutofilter());
      assertInstanceOf(ExcelTableStyleSnapshot.Named.class, snapshot.style());
    }
  }
}
