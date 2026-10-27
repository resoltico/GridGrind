package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetCopyPosition;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
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

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      seedCalculatedTableSheet(workbook, "Ledger", "OpsCloseTable");

      workbook.sheets().copySheet("Ledger", "Ledger Copy", new SheetCopyPosition.AppendAtEnd());

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
      assertDoesNotThrow(() -> workbook.sheets().renameSheet("Ledger Copy", "Ledger Final"));

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

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
    }

    try (XssfWorkbook workbook =
        XssfWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
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
    assertEquals(Optional.empty(), TableAdapter.calculatedColumnFormula(missingFormula));

    CTTableColumn blankFormula = CTTableColumn.Factory.newInstance();
    blankFormula.addNewCalculatedColumnFormula().setStringValue(" ");
    assertEquals(Optional.empty(), TableAdapter.calculatedColumnFormula(blankFormula));

    CTTableColumn valuedFormula = CTTableColumn.Factory.newInstance();
    valuedFormula.addNewCalculatedColumnFormula().setStringValue("1+1");
    assertEquals(Optional.of("1+1"), TableAdapter.calculatedColumnFormula(valuedFormula));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      CreationHelper creationHelper = workbook.getCreationHelper();
      var drawing = sheet.createDrawingPatriarch();
      var row = sheet.createRow(0);

      var plain = row.createCell(0);
      assertFalse(TableAdapter.mustPreserveCellShell(plain));

      var commented = row.createCell(1);
      var comment = drawing.createCellComment(drawing.createAnchor(0, 0, 0, 0, 1, 0, 2, 2));
      comment.setString(creationHelper.createRichTextString("keep"));
      commented.setCellComment(comment);
      assertTrue(TableAdapter.mustPreserveCellShell(commented));

      var hyperlinked = row.createCell(2);
      hyperlinked.setHyperlink(creationHelper.createHyperlink(HyperlinkType.URL));
      assertTrue(TableAdapter.mustPreserveCellShell(hyperlinked));

      var styled = row.createCell(3);
      styled.setCellStyle(workbook.createCellStyle());
      assertTrue(TableAdapter.mustPreserveCellShell(styled));
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

      TableAdapter.canonicalizeSheet(sheet);

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

      assertDoesNotThrow(() -> TableAdapter.canonicalizeSheet(invalidRefSheet));
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

      TableAdapter.canonicalizeSheet(noBodySheet);

      assertEquals("1+1", noBodySheet.getRow(1).getCell(1).getCellFormula());
    }
  }

  private static void seedCalculatedTableSheet(
      XssfWorkbook workbook, String sheetName, String tableName) throws IOException {
    XssfSheet sheet = workbook.getOrCreateSheet(sheetName);
    sheet.cells().setCell("A1", new CellValue.Text("Task"));
    sheet.cells().setCell("B1", new CellValue.Text("Actual"));
    sheet.cells().setCell("C1", new CellValue.Text("Budget"));
    sheet.cells().setCell("D1", new CellValue.Text("Delta"));
    sheet.cells().setCell("A2", new CellValue.Text("One"));
    sheet.cells().setCell("B2", new CellValue.NumberValue(10));
    sheet.cells().setCell("C2", new CellValue.NumberValue(7));
    sheet.cells().setCell("A3", new CellValue.Text("Two"));
    sheet.cells().setCell("B3", new CellValue.NumberValue(5));
    sheet.cells().setCell("C3", new CellValue.NumberValue(6));
    sheet.cells().setCell("A4", new CellValue.Text("Total"));
    sheet.cells().setCell("B4", new CellValue.Blank());
    sheet.cells().setCell("C4", new CellValue.Blank());
    sheet.cells().setCell("D4", new CellValue.Blank());
    new TableAdapter()
        .setTable(
            workbook,
            new ResolvedTable(
                new TableName(tableName),
                new SheetName(sheetName),
                new RangeRef("A1:D4"),
                true,
                true,
                new TableStyleSpec.None(),
                "",
                false,
                false,
                false,
                "",
                "",
                "",
                List.of(
                    new TableColumnSpec(1, "", "", "sum", ""),
                    new TableColumnSpec(2, "", "", "sum", ""),
                    new TableColumnSpec(3, "", "", "sum", "[@Actual]-[@Budget]"))));
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
