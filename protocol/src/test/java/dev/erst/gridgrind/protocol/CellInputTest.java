package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelCellValue;
import java.time.LocalDate;
import java.time.LocalDateTime;
import org.junit.jupiter.api.Test;

/** Tests for CellInput record construction and ExcelCellValue conversion. */
class CellInputTest {
  @Test
  void convertsAllSupportedInputTypesToExcelValues() {
    assertInstanceOf(ExcelCellValue.BlankValue.class, new CellInput.Blank().toExcelCellValue());

    ExcelCellValue.TextValue textValue =
        assertInstanceOf(
            ExcelCellValue.TextValue.class, new CellInput.Text("Budget").toExcelCellValue());
    assertEquals("Budget", textValue.value());

    ExcelCellValue.NumberValue numberValue =
        assertInstanceOf(
            ExcelCellValue.NumberValue.class, new CellInput.Numeric(42.5).toExcelCellValue());
    assertEquals(42.5, numberValue.value());

    ExcelCellValue.BooleanValue booleanValue =
        assertInstanceOf(
            ExcelCellValue.BooleanValue.class, new CellInput.BooleanValue(true).toExcelCellValue());
    assertTrue(booleanValue.value());

    ExcelCellValue.FormulaValue formulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            new CellInput.Formula("SUM(B2:B4)").toExcelCellValue());
    assertEquals("SUM(B2:B4)", formulaValue.expression());

    ExcelCellValue.DateValue dateValue =
        assertInstanceOf(
            ExcelCellValue.DateValue.class,
            new CellInput.Date(LocalDate.of(2026, 3, 23)).toExcelCellValue());
    assertEquals(LocalDate.of(2026, 3, 23), dateValue.value());

    ExcelCellValue.DateTimeValue dateTimeValue =
        assertInstanceOf(
            ExcelCellValue.DateTimeValue.class,
            new CellInput.DateTime(LocalDateTime.of(2026, 3, 23, 10, 15, 30)).toExcelCellValue());
    assertEquals(LocalDateTime.of(2026, 3, 23, 10, 15, 30), dateTimeValue.value());
  }

  @Test
  void validatesTypedInputRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Text(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Numeric(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.BooleanValue(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Formula(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Date(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.DateTime(null));
  }
}
