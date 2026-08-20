package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import org.apache.poi.ss.usermodel.FormulaError;
import org.junit.jupiter.api.Test;

/** Tests for ExcelCellValue sealed interface record construction. */
class ExcelCellValueTest {
  @Test
  void createsAllSupportedExcelCellValues() {
    assertInstanceOf(ExcelCellValue.BlankValue.class, ExcelCellValue.blank());
    assertEquals(
        "Budget",
        assertInstanceOf(ExcelCellValue.TextValue.class, ExcelCellValue.text("Budget")).value());
    assertEquals(
        12.5,
        assertInstanceOf(ExcelCellValue.NumberValue.class, ExcelCellValue.number(12.5)).value());
    assertTrue(
        assertInstanceOf(ExcelCellValue.BooleanValue.class, ExcelCellValue.bool(true)).value());
    assertEquals(
        "#REF!",
        assertInstanceOf(ExcelCellValue.ErrorValue.class, ExcelCellValue.error("#REF!")).value());
    assertEquals(
        LocalDate.of(2026, 3, 23),
        assertInstanceOf(
                ExcelCellValue.DateValue.class, ExcelCellValue.date(LocalDate.of(2026, 3, 23)))
            .value());
    assertEquals(
        LocalDateTime.of(2026, 3, 23, 8, 30),
        assertInstanceOf(
                ExcelCellValue.DateTimeValue.class,
                ExcelCellValue.dateTime(LocalDateTime.of(2026, 3, 23, 8, 30)))
            .value());
    assertEquals(
        "SUM(A1:A3)",
        assertInstanceOf(ExcelCellValue.FormulaValue.class, ExcelCellValue.formula("SUM(A1:A3)"))
            .expression());
    assertEquals(
        "LAMBDA(x,x+1)(A1)",
        assertInstanceOf(
                ExcelCellValue.RawFormulaValue.class,
                ExcelCellValue.rawFormula("LAMBDA(x,x+1)(A1)"))
            .expression());
  }

  @Test
  void validatesNonNullAndNonBlankValueRequirements() {
    assertThrows(NullPointerException.class, () -> ExcelCellValue.text(null));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.error(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellValue.error(" "));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellValue.error("#CIRCULAR_REF!"));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelCellValue.error("#FUNCTION_NOT_IMPLEMENTED!"));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellValue.error("~CIRCULAR~REF~"));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.date(null));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.dateTime(null));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.formula(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellValue.formula(" "));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.rawFormula(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellValue.rawFormula(" "));
  }

  @Test
  void rejectsPoiFormulaErrorsThatHaveNoPublishedGridGrindWireLiteral() {
    IllegalArgumentException unsupported =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelCellErrorLiteralSupport.toReportedWireValue(FormulaError._NO_ERROR));

    assertEquals("_NO_ERROR is not a publishable cell error", unsupported.getMessage());
  }

  @Test
  void writesOpaqueFormulaValuesThroughTheSheetMutationBoundary() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook
          .getOrCreateSheet("Budget")
          .cells()
          .setCell("A1", ExcelCellValue.rawFormula("LAMBDA(x,x+1)(A1)"));

      assertEquals(
          "LAMBDA(x,x+1)(A1)",
          workbook.xssfWorkbook().getSheet("Budget").getRow(0).getCell(0).getCellFormula());
    }
  }
}
