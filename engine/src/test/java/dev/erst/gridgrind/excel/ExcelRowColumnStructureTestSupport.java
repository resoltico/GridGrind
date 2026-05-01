package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import java.io.IOException;
import java.util.List;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.function.Executable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;

/** Shared helpers for row and column structure coverage slices. */
class ExcelRowColumnStructureTestSupport {
  final ExcelRowColumnStructureController controller = new ExcelRowColumnStructureController();

  static void seedSupportedScenario(
      XSSFWorkbook workbook, XSSFSheet sheet, boolean includeFormulas) {
    setString(sheet, "A1", "Item");
    setString(sheet, "B1", "Value");
    setString(sheet, "C1", "Status");
    setString(sheet, "D1", "Notes");
    setString(sheet, "E1", "Aux");
    setString(sheet, "F1", "Merge");
    setString(sheet, "G1", "Formula");

    setString(sheet, "A2", "Hosting");
    setNumeric(sheet, "B2", 42);
    setString(sheet, "C2", "Open");
    setString(sheet, "D2", "Alpha");
    setString(sheet, "A3", "Support");
    setNumeric(sheet, "B3", 84);
    setString(sheet, "C3", "Closed");
    setString(sheet, "D3", "Beta");
    setString(sheet, "A4", "Ops");
    setNumeric(sheet, "B4", 168);
    setString(sheet, "C4", "Open");
    setString(sheet, "D4", "Gamma");
    setString(sheet, "A5", "Tail");
    setNumeric(sheet, "B5", 7);

    if (includeFormulas) {
      sheet.getRow(1).createCell(6).setCellFormula("SUM(B2:B4)");
      sheet.getRow(2).createCell(6).setCellFormula("A2&B2");
      setString(sheet, "G12", "Anchor");
      sheet.getRow(11).createCell(7).setCellFormula("SUM(B2:B4)");
      sheet.getRow(11).createCell(8).setCellFormula("A2&B2");
    }

    Name name = workbook.createName();
    name.setNameName("BudgetValues");
    name.setRefersToFormula("Budget!$B$2:$B$4");

    sheet.addMergedRegion(CellRangeAddress.valueOf("E2:F3"));

    SheetConditionalFormatting conditionalFormatting = sheet.getSheetConditionalFormatting();
    XSSFConditionalFormattingRule rule =
        (XSSFConditionalFormattingRule)
            conditionalFormatting.createConditionalFormattingRule("B2>100");
    conditionalFormatting.addConditionalFormatting(
        new CellRangeAddress[] {CellRangeAddress.valueOf("B2:B5")}, rule);

    Cell cellWithLink = getOrCreateCell(sheet, "D5");
    XSSFHyperlink hyperlink =
        (XSSFHyperlink) workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
    hyperlink.setAddress("https://example.com/report");
    cellWithLink.setHyperlink(hyperlink);

    XSSFComment comment =
        (XSSFComment)
            sheet
                .createDrawingPatriarch()
                .createCellComment(workbook.getCreationHelper().createClientAnchor());
    comment.setString(workbook.getCreationHelper().createRichTextString("Review"));
    cellWithLink.setCellComment(comment);
  }

  static void seedTable(XSSFSheet sheet, XSSFWorkbook workbook) {
    seedTable(sheet, workbook, "BudgetTable");
  }

  static void seedTable(XSSFSheet sheet, XSSFWorkbook workbook, String tableName) {
    setString(sheet, "A1", "Item");
    setString(sheet, "B1", "Value");
    setString(sheet, "A2", "Hosting");
    setNumeric(sheet, "B2", 42);
    setString(sheet, "A3", "Support");
    setNumeric(sheet, "B3", 84);
    XSSFTable table =
        sheet.createTable(new AreaReference("A1:B3", workbook.getSpreadsheetVersion()));
    table.setName(tableName);
    table.setDisplayName(tableName);
  }

  static void seedNamedRange(XSSFWorkbook workbook, String nameName, String refersToFormula) {
    Name name = workbook.createName();
    name.setNameName(nameName);
    name.setRefersToFormula(refersToFormula);
  }

  static void seedBlankNamedRange(XSSFWorkbook workbook) {
    var definedNames =
        workbook.getCTWorkbook().isSetDefinedNames()
            ? workbook.getCTWorkbook().getDefinedNames()
            : workbook.getCTWorkbook().addNewDefinedNames();
    var blankName = definedNames.addNewDefinedName();
    blankName.setName("BlankBudget");
    blankName.setStringValue(" ");
  }

  static void seedSheetAutofilter(XSSFSheet sheet) {
    setString(sheet, "A1", "Item");
    setString(sheet, "B1", "Value");
    setString(sheet, "A2", "Hosting");
    setNumeric(sheet, "B2", 42);
    sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B2"));
  }

  static void seedDataValidation(XSSFSheet sheet) {
    setString(sheet, "A1", "Status");
    setString(sheet, "A2", "Open");
    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint =
        helper.createExplicitListConstraint(new String[] {"Open", "Closed"});
    DataValidation validation =
        helper.createValidation(constraint, new CellRangeAddressList(1, 3, 0, 0));
    sheet.addValidationData(validation);
  }

  static List<String> dataValidationRanges(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    List<String> ranges = new java.util.ArrayList<>();
    for (var validation : sheet.getCTWorksheet().getDataValidations().getDataValidationArray()) {
      ranges.addAll(ExcelSqrefSupport.normalizedSqref(validation.getSqref()));
    }
    return List.copyOf(ranges);
  }

  IllegalArgumentException unsupportedStructure(Executable executable) {
    return assertThrows(IllegalArgumentException.class, executable);
  }

  static List<String> hyperlinkAddresses(XSSFSheet sheet) {
    return sheet.getHyperlinkList().stream()
        .map(hyperlink -> hyperlink.getCellRef())
        .sorted()
        .toList();
  }

  static List<String> commentAddresses(XSSFSheet sheet) {
    List<String> addresses = new java.util.ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellComment() != null) {
          addresses.add(cell.getAddress().formatAsString());
        }
      }
    }
    return List.copyOf(addresses);
  }

  static void setString(XSSFSheet sheet, String address, String value) {
    getOrCreateCell(sheet, address).setCellValue(value);
  }

  static void setNumeric(XSSFSheet sheet, String address, double value) {
    getOrCreateCell(sheet, address).setCellValue(value);
  }

  static Cell getOrCreateCell(XSSFSheet sheet, String address) {
    CellReference reference = new CellReference(address);
    Row row = sheet.getRow(reference.getRow());
    if (row == null) {
      row = sheet.createRow(reference.getRow());
    }
    Cell cell = row.getCell(reference.getCol());
    if (cell == null) {
      cell = row.createCell(reference.getCol());
    }
    return cell;
  }

  static List<ColumnGroupingMutation> columnGroupingMutations() {
    List<ColumnGroupingMutation> mutations = new java.util.ArrayList<>();
    for (int firstColumnIndex = 0; firstColumnIndex <= 3; firstColumnIndex++) {
      for (int lastColumnIndex = firstColumnIndex; lastColumnIndex <= 3; lastColumnIndex++) {
        mutations.add(new ColumnGroupingMutation("group", firstColumnIndex, lastColumnIndex, true));
        mutations.add(
            new ColumnGroupingMutation("group", firstColumnIndex, lastColumnIndex, false));
        mutations.add(
            new ColumnGroupingMutation("ungroup", firstColumnIndex, lastColumnIndex, false));
      }
    }
    return List.copyOf(mutations);
  }

  void assertColumnOutlineLevels(XSSFSheet sheet, int first, int second, int third) {
    List<WorkbookSheetResult.ColumnLayout> columns = controller.columnLayouts(sheet);

    assertEquals(first, columns.get(1).outlineLevel(), "columns=" + columns);
    assertEquals(second, columns.get(2).outlineLevel(), "columns=" + columns);
    assertEquals(third, columns.get(3).outlineLevel(), "columns=" + columns);
  }

  static void assertCanonicalColumnDefinitions(XSSFSheet sheet) {
    assertEquals(1, sheet.getCTWorksheet().sizeOfColsArray());

    boolean[] seenColumns = new boolean[ExcelColumnSpan.MAX_COLUMN_INDEX + 1];
    for (CTCol col : sheet.getCTWorksheet().getColsArray(0).getColList()) {
      assertEquals(col.getMin(), col.getMax(), "column definitions=" + sheet.getCTWorksheet());
      for (int columnIndex = (int) col.getMin() - 1;
          columnIndex <= (int) col.getMax() - 1;
          columnIndex++) {
        assertFalse(
            seenColumns[columnIndex], "column definitions overlap for index " + columnIndex);
        seenColumns[columnIndex] = true;
      }
    }
  }

  static CTCol addRawColumnDefinition(XSSFSheet sheet, int firstColumnIndex, int lastColumnIndex) {
    CTCols cols =
        sheet.getCTWorksheet().sizeOfColsArray() == 0
            ? sheet.getCTWorksheet().addNewCols()
            : sheet.getCTWorksheet().getColsArray(0);
    CTCol definition = cols.addNewCol();
    definition.setMin(firstColumnIndex + 1L);
    definition.setMax(lastColumnIndex + 1L);
    return definition;
  }

  record ColumnGroupingMutation(
      String operation, int firstColumnIndex, int lastColumnIndex, boolean collapsed) {
    void apply(ExcelRowColumnStructureController controller, XSSFSheet sheet) {
      ExcelColumnSpan columns = new ExcelColumnSpan(firstColumnIndex, lastColumnIndex);
      if ("group".equals(operation)) {
        controller.groupColumns(sheet, columns, collapsed);
        return;
      }
      controller.ungroupColumns(sheet, columns);
    }

    @Override
    public String toString() {
      return operation
          + "("
          + firstColumnIndex
          + ","
          + lastColumnIndex
          + ("group".equals(operation) ? "," + collapsed : "")
          + ")";
    }
  }

  void assertUnexpectedRuntimeFailureAbsent(
      ColumnGroupingMutation first, ColumnGroupingMutation second) {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Budget");
      try {
        first.apply(controller, sheet);
        second.apply(controller, sheet);
      } catch (IllegalArgumentException expected) {
        // Invalid outline edits are allowed; only unexpected runtime failures should fail.
      } catch (RuntimeException unexpected) {
        fail(first + " -> " + second + " threw unexpected runtime failure", unexpected);
      }
    } catch (IOException ioException) {
      fail("Closing workbook should not fail during grouping regression checks", ioException);
    }
  }

  /** Minimal POI Name stub for direct defined-name predicate tests. */
  static final class DefinedNameStub implements Name {
    final String refersToFormula;
    final int sheetIndex;

    DefinedNameStub(String refersToFormula, int sheetIndex) {
      this.refersToFormula = refersToFormula;
      this.sheetIndex = sheetIndex;
    }

    @Override
    public String getSheetName() {
      return sheetIndex < 0 ? null : "Budget";
    }

    @Override
    public String getNameName() {
      return "TestName";
    }

    @Override
    public void setNameName(String name) {
      throw new UnsupportedOperationException();
    }

    @Override
    public String getRefersToFormula() {
      return refersToFormula;
    }

    @Override
    public void setRefersToFormula(String formulaText) {
      throw new UnsupportedOperationException();
    }

    @Override
    public boolean isFunctionName() {
      return false;
    }

    @Override
    public boolean isDeleted() {
      return false;
    }

    @Override
    public boolean isHidden() {
      return false;
    }

    @Override
    public void setSheetIndex(int index) {
      throw new UnsupportedOperationException();
    }

    @Override
    public int getSheetIndex() {
      return sheetIndex;
    }

    @Override
    public String getComment() {
      return "";
    }

    @Override
    public void setComment(String comment) {
      throw new UnsupportedOperationException();
    }

    @Override
    public void setFunction(boolean value) {
      throw new UnsupportedOperationException();
    }
  }
}
