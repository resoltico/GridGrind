package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;

/** Regressions for copied-sheet calculated-column canonicalization and rename stability. */
class ExcelTableCalculatedColumnCanonicalizerTest {
  @Test
  void copySheetCanonicalizesMaterializedCalculatedColumnCellsAndRenameWorks() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-copy-table-canonical-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      seedCalculatedTableSheet(workbook, "Ledger", "OpsCloseTable");

      workbook.copySheet("Ledger", "Ledger Copy", new ExcelSheetCopyPosition.AppendAtEnd());

      assertNoFormulaCell(workbook.xssfWorkbook().getSheet("Ledger Copy"), "D2");
      assertNoFormulaCell(workbook.xssfWorkbook().getSheet("Ledger Copy"), "D3");
      assertEquals(
          "SUBTOTAL(109,OpsCloseTable_Copy2[Actual])",
          workbook.xssfWorkbook().getSheet("Ledger Copy").getRow(3).getCell(1).getCellFormula());
      assertEquals(
          "SUBTOTAL(109,OpsCloseTable_Copy2[Budget])",
          workbook.xssfWorkbook().getSheet("Ledger Copy").getRow(3).getCell(2).getCellFormula());
      assertEquals(
          "SUBTOTAL(109,OpsCloseTable_Copy2[Delta])",
          workbook.xssfWorkbook().getSheet("Ledger Copy").getRow(3).getCell(3).getCellFormula());
      assertDoesNotThrow(() -> workbook.renameSheet("Ledger Copy", "Ledger Final"));

      XSSFTable copiedTable =
          workbook.xssfWorkbook().getSheet("Ledger Final").getTables().getFirst();
      assertEquals(
          "[@Actual]-[@Budget]",
          copiedTable
              .getCTTable()
              .getTableColumns()
              .getTableColumnArray(3)
              .getCalculatedColumnFormula()
              .getStringValue());

      workbook.save(workbookPath);
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      XSSFSheet renamedSheet = workbook.xssfWorkbook().getSheet("Ledger Final");
      assertNoFormulaCell(renamedSheet, "D2");
      assertNoFormulaCell(renamedSheet, "D3");
      assertEquals(
          "[@Actual]-[@Budget]",
          renamedSheet
              .getTables()
              .getFirst()
              .getCTTable()
              .getTableColumns()
              .getTableColumnArray(3)
              .getCalculatedColumnFormula()
              .getStringValue());
    }
  }

  @Test
  void helperMethodsCoverBlankFormulaAndPreserveShellCases() throws IOException {
    CTTableColumn missingFormula = CTTableColumn.Factory.newInstance();
    assertNull(ExcelTableCalculatedColumnCanonicalizer.calculatedColumnFormula(missingFormula));

    CTTableColumn blankFormula = CTTableColumn.Factory.newInstance();
    blankFormula.addNewCalculatedColumnFormula().setStringValue(" ");
    assertNull(ExcelTableCalculatedColumnCanonicalizer.calculatedColumnFormula(blankFormula));

    CTTableColumn valuedFormula = CTTableColumn.Factory.newInstance();
    valuedFormula.addNewCalculatedColumnFormula().setStringValue("1+1");
    assertEquals(
        "1+1", ExcelTableCalculatedColumnCanonicalizer.calculatedColumnFormula(valuedFormula));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      CreationHelper creationHelper = workbook.getCreationHelper();
      var drawing = sheet.createDrawingPatriarch();
      var row = sheet.createRow(0);

      var plain = row.createCell(0);
      assertFalse(ExcelTableCalculatedColumnCanonicalizer.mustPreserveCellShell(plain));

      var commented = row.createCell(1);
      var comment = drawing.createCellComment(drawing.createAnchor(0, 0, 0, 0, 1, 0, 2, 2));
      comment.setString(creationHelper.createRichTextString("keep"));
      commented.setCellComment(comment);
      assertTrue(ExcelTableCalculatedColumnCanonicalizer.mustPreserveCellShell(commented));

      var hyperlinked = row.createCell(2);
      hyperlinked.setHyperlink(creationHelper.createHyperlink(HyperlinkType.URL));
      assertTrue(ExcelTableCalculatedColumnCanonicalizer.mustPreserveCellShell(hyperlinked));

      var styled = row.createCell(3);
      styled.setCellStyle(workbook.createCellStyle());
      assertTrue(ExcelTableCalculatedColumnCanonicalizer.mustPreserveCellShell(styled));
    }
  }

  @Test
  void canonicalizeSheetHandlesInvalidRefsBodyGapsAndRemovableCells() throws IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      seedCanonicalizerWorksheet(sheet);

      XSSFTable table = sheet.createTable(new AreaReference("A1:B7", SpreadsheetVersion.EXCEL2007));
      table
          .getCTTable()
          .getTableColumns()
          .getTableColumnArray(1)
          .addNewCalculatedColumnFormula()
          .setStringValue("1+1");

      ExcelTableCalculatedColumnCanonicalizer.canonicalizeSheet(sheet);

      assertEquals(CellType.BLANK, sheet.getRow(1).getCell(1).getCellType());
      assertNull(sheet.getRow(2).getCell(1));
      assertEquals(CellType.STRING, sheet.getRow(3).getCell(1).getCellType());
      assertEquals("2+2", sheet.getRow(4).getCell(1).getCellFormula());
      assertNull(sheet.getRow(5));
      assertNull(sheet.getRow(6).getCell(1));
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet invalidRefSheet = workbook.createSheet("InvalidRef");
      invalidRefSheet.createRow(0).createCell(0).setCellValue("Name");
      invalidRefSheet.getRow(0).createCell(1).setCellValue("Calc");
      invalidRefSheet.createRow(1).createCell(0).setCellValue("Ada");
      invalidRefSheet.getRow(1).createCell(1).setCellFormula("1+1");
      XSSFTable invalidRefTable =
          invalidRefSheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
      invalidRefTable.getCTTable().setRef("");
      invalidRefTable
          .getCTTable()
          .getTableColumns()
          .getTableColumnArray(1)
          .addNewCalculatedColumnFormula()
          .setStringValue("1+1");

      assertDoesNotThrow(
          () -> ExcelTableCalculatedColumnCanonicalizer.canonicalizeSheet(invalidRefSheet));
      assertEquals("1+1", invalidRefSheet.getRow(1).getCell(1).getCellFormula());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet noBodySheet = workbook.createSheet("NoBody");
      noBodySheet.createRow(0).createCell(0).setCellValue("Name");
      noBodySheet.getRow(0).createCell(1).setCellValue("Calc");
      noBodySheet.createRow(1).createCell(0).setCellValue("Total");
      noBodySheet.getRow(1).createCell(1).setCellFormula("1+1");
      XSSFTable noBodyTable =
          noBodySheet.createTable(new AreaReference("A1:B2", SpreadsheetVersion.EXCEL2007));
      noBodyTable.getCTTable().setTotalsRowCount(1);
      noBodyTable
          .getCTTable()
          .getTableColumns()
          .getTableColumnArray(1)
          .addNewCalculatedColumnFormula()
          .setStringValue(" ");

      ExcelTableCalculatedColumnCanonicalizer.canonicalizeSheet(noBodySheet);

      assertEquals("1+1", noBodySheet.getRow(1).getCell(1).getCellFormula());
    }
  }

  private static void seedCalculatedTableSheet(
      ExcelWorkbook workbook, String sheetName, String tableName) throws IOException {
    ExcelSheet sheet = workbook.getOrCreateSheet(sheetName);
    sheet.setCell("A1", ExcelCellValue.text("Task"));
    sheet.setCell("B1", ExcelCellValue.text("Actual"));
    sheet.setCell("C1", ExcelCellValue.text("Budget"));
    sheet.setCell("D1", ExcelCellValue.text("Delta"));
    sheet.setCell("A2", ExcelCellValue.text("One"));
    sheet.setCell("B2", ExcelCellValue.number(10));
    sheet.setCell("C2", ExcelCellValue.number(7));
    sheet.setCell("A3", ExcelCellValue.text("Two"));
    sheet.setCell("B3", ExcelCellValue.number(5));
    sheet.setCell("C3", ExcelCellValue.number(6));
    sheet.setCell("A4", ExcelCellValue.text("Total"));
    sheet.setCell("B4", ExcelCellValue.blank());
    sheet.setCell("C4", ExcelCellValue.blank());
    sheet.setCell("D4", ExcelCellValue.blank());
    workbook.setTable(
        new ExcelTableDefinition(
            tableName,
            sheetName,
            "A1:D4",
            true,
            true,
            new ExcelTableStyle.None(),
            "",
            false,
            false,
            false,
            "",
            "",
            "",
            List.of(
                new ExcelTableColumnDefinition(1, "", "", "sum", ""),
                new ExcelTableColumnDefinition(2, "", "", "sum", ""),
                new ExcelTableColumnDefinition(3, "", "", "sum", "[@Actual]-[@Budget]"))));
  }

  private static void seedCanonicalizerWorksheet(XSSFSheet sheet) {
    sheet.createRow(0).createCell(0).setCellValue("Name");
    sheet.getRow(0).createCell(1).setCellValue("Calc");

    CreationHelper creationHelper = sheet.getWorkbook().getCreationHelper();
    var drawing = sheet.createDrawingPatriarch();

    sheet.createRow(1).createCell(0).setCellValue("KeepShell");
    sheet.getRow(1).createCell(1).setCellFormula("1+1");
    var comment = drawing.createCellComment(drawing.createAnchor(0, 0, 0, 0, 1, 1, 2, 3));
    comment.setString(creationHelper.createRichTextString("keep"));
    sheet.getRow(1).getCell(1).setCellComment(comment);

    sheet.createRow(2).createCell(0).setCellValue("MissingCell");

    sheet.createRow(3).createCell(0).setCellValue("Literal");
    sheet.getRow(3).createCell(1).setCellValue("literal");

    sheet.createRow(4).createCell(0).setCellValue("Mismatch");
    sheet.getRow(4).createCell(1).setCellFormula("2+2");

    sheet.createRow(6).createCell(0).setCellValue("RemoveCell");
    sheet.getRow(6).createCell(1).setCellFormula("1+1");
  }

  private static void assertNoFormulaCell(XSSFSheet sheet, String address) {
    var reference = new CellReference(address);
    var row = sheet.getRow(reference.getRow());
    if (row == null) {
      return;
    }
    var cell = row.getCell(reference.getCol());
    if (cell == null) {
      return;
    }
    assertNotEquals(
        CellType.FORMULA, cell.getCellType(), address + " must not remain a formula cell");
  }
}
