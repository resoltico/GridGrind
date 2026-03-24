package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
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
  }

  @Test
  void validatesNonNullAndNonBlankValueRequirements() {
    assertThrows(NullPointerException.class, () -> ExcelCellValue.text(null));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.date(null));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.dateTime(null));
    assertThrows(NullPointerException.class, () -> ExcelCellValue.formula(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellValue.formula(" "));
  }
}
