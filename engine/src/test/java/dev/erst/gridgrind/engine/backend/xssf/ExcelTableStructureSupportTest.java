package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.engine.poi.CellBounds;
import dev.erst.gridgrind.engine.poi.SheetStructure;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.lang.reflect.Proxy;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Tests for package-private table-structure helpers shared across authoring and analysis. */
class ExcelTableStructureSupportTest {
  @Test
  void requireSupportedTableShapeValidatesDataAndTotalsLayouts() {
    assertDoesNotThrow(
        () -> TableAdapter.requireSupportedTableShape(CellBounds.parse("A1:B2"), false));
    assertDoesNotThrow(
        () -> TableAdapter.requireSupportedTableShape(CellBounds.parse("A1:B3"), true));
    assertThrows(
        IllegalArgumentException.class,
        () -> TableAdapter.requireSupportedTableShape(CellBounds.parse("A1:B1"), false));
    assertThrows(
        IllegalArgumentException.class,
        () -> TableAdapter.requireSupportedTableShape(CellBounds.parse("A1:B2"), true));
  }

  @Test
  void headerNamesAndHeaderTextNormalizeSupportedCellKinds() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      assertEquals(List.of("", ""), SheetStructure.headerNames(sheet, CellBounds.parse("A1:B1")));

      var row = sheet.createRow(0);
      row.createCell(1).setBlank();
      row.createCell(2).setCellValue(" Owner ");
      row.createCell(3).setCellFormula("A1");
      row.createCell(4).setCellValue(12.0);
      row.createCell(5).setCellValue(true);

      assertEquals("", SheetStructure.headerText(null));
      assertEquals("", SheetStructure.headerText(cellStub(CellType._NONE, "", "", "")));
      assertEquals("", SheetStructure.headerText(cellStub(CellType.FORMULA, "", "", "")));
      assertEquals("", SheetStructure.headerText(row.getCell(1)));
      assertEquals("Owner", SheetStructure.headerText(row.getCell(2)));
      assertEquals("A1", SheetStructure.headerText(row.getCell(3)));
      assertEquals("12.0", SheetStructure.headerText(row.getCell(4)));
      assertEquals("TRUE", SheetStructure.headerText(row.getCell(5)));
      assertEquals(
          List.of("", "Owner", "A1", "12.0", "TRUE"),
          SheetStructure.headerNames(sheet, CellBounds.parse("B1:F1")));
    }
  }

  @Test
  void applyAutofilterAddsAndReusesExistingAutofilterMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFTable table = createTable(workbook, "Ops", "A1:B4");

      TableAdapter.applyAutofilter(table, CellBounds.parse("A1:B4"), false);
      assertEquals("A1:B4", table.getCTTable().getAutoFilter().getRef());

      TableAdapter.applyAutofilter(table, CellBounds.parse("A1:B4"), true);
      assertEquals("A1:B3", table.getCTTable().getAutoFilter().getRef());
    }
  }

  @Test
  void applyStyleAddsReusesAndClearsStyleMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFTable table = createTable(workbook, "Ops", "A1:B3");

      TableAdapter.applyStyle(
          table, new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false));
      assertEquals("TableStyleMedium2", table.getCTTable().getTableStyleInfo().getName());
      assertTrue(table.getCTTable().isSetTableStyleInfo());

      TableAdapter.applyStyle(
          table, new TableStyleSpec.Named("TableStyleMedium2", true, true, false, true));
      assertTrue(table.getCTTable().getTableStyleInfo().getShowFirstColumn());
      assertTrue(table.getCTTable().getTableStyleInfo().getShowLastColumn());
      assertFalse(table.getCTTable().getTableStyleInfo().getShowRowStripes());
      assertTrue(table.getCTTable().getTableStyleInfo().getShowColumnStripes());

      TableAdapter.applyStyle(table, new TableStyleSpec.None());
      assertFalse(table.getCTTable().isSetTableStyleInfo());
    }
  }

  @Test
  void expectedAutofilterRangeTextHandlesValidTotalsAndInvalidRanges() {
    assertEquals(
        "A1:B3",
        TableAdapter.expectedAutofilterRangeText(
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B4"),
                1,
                1,
                List.of("Owner", "Task"),
                new TableStyleReport.None(),
                true)));
    assertEquals(
        "A1:B3",
        TableAdapter.expectedAutofilterRangeText(
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                1,
                0,
                List.of("Owner", "Task"),
                new TableStyleReport.None(),
                true)));
    assertEquals(
        "A0:B3",
        TableAdapter.expectedAutofilterRangeText(
            new TableEntryReport(
                new dev.erst.gridgrind.core.vocabulary.TableLabel("Queue"),
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                new dev.erst.gridgrind.core.vocabulary.RangeLabel("A0:B3"),
                1,
                0,
                List.of("Owner", "Task"),
                new TableStyleReport.None(),
                true)));
  }

  private static XSSFTable createTable(XSSFWorkbook workbook, String sheetName, String range) {
    XSSFSheet sheet = workbook.createSheet(sheetName);
    sheet.createRow(0).createCell(0).setCellValue("Owner");
    sheet.getRow(0).createCell(1).setCellValue("Task");
    sheet.createRow(1).createCell(0).setCellValue("Ada");
    sheet.getRow(1).createCell(1).setCellValue("Queue");
    sheet.createRow(2).createCell(0).setCellValue("Lin");
    sheet.getRow(2).createCell(1).setCellValue("Pack");
    sheet.createRow(3).createCell(0).setCellValue("Totals");
    sheet.getRow(3).createCell(1).setCellValue("Done");
    return sheet.createTable(new AreaReference(range, SpreadsheetVersion.EXCEL2007));
  }

  private static org.apache.poi.ss.usermodel.Cell cellStub(
      CellType cellType, String stringValue, String formulaValue, String displayValue) {
    return (org.apache.poi.ss.usermodel.Cell)
        Proxy.newProxyInstance(
            Objects.requireNonNull(Thread.currentThread().getContextClassLoader()),
            new Class<?>[] {org.apache.poi.ss.usermodel.Cell.class},
            (proxy, method, args) ->
                switch (method.getName()) {
                  case "getCellType" -> cellType;
                  case "getStringCellValue" -> stringValue;
                  case "getCellFormula" -> formulaValue;
                  case "toString" -> displayValue;
                  case "hashCode" -> System.identityHashCode(proxy);
                  case "equals" -> false;
                  default -> throw new UnsupportedOperationException(method.getName());
                });
  }
}
