package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.protocol.dto.CellInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import org.junit.jupiter.api.Test;

/** Tests for CellInput record construction and ExcelCellValue conversion. */
class CellInputTest {
  @Test
  void convertsAllSupportedInputTypesToExcelValues() {
    assertInstanceOf(
        ExcelCellValue.BlankValue.class,
        DefaultGridGrindRequestExecutor.toExcelCellValue(new CellInput.Blank()));

    ExcelCellValue.TextValue textValue =
        assertInstanceOf(
            ExcelCellValue.TextValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(new CellInput.Text("Budget")));
    assertEquals("Budget", textValue.value());

    ExcelCellValue.NumberValue numberValue =
        assertInstanceOf(
            ExcelCellValue.NumberValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(new CellInput.Numeric(42.5)));
    assertEquals(42.5, numberValue.value());

    ExcelCellValue.BooleanValue booleanValue =
        assertInstanceOf(
            ExcelCellValue.BooleanValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(new CellInput.BooleanValue(true)));
    assertTrue(booleanValue.value());

    ExcelCellValue.FormulaValue formulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(new CellInput.Formula("SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", formulaValue.expression());

    // Leading = is stripped automatically so callers can use Excel-native syntax
    ExcelCellValue.FormulaValue strippedFormulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(new CellInput.Formula("=SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", strippedFormulaValue.expression());

    ExcelCellValue.DateValue dateValue =
        assertInstanceOf(
            ExcelCellValue.DateValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(
                new CellInput.Date(LocalDate.of(2026, 3, 23))));
    assertEquals(LocalDate.of(2026, 3, 23), dateValue.value());

    ExcelCellValue.DateTimeValue dateTimeValue =
        assertInstanceOf(
            ExcelCellValue.DateTimeValue.class,
            DefaultGridGrindRequestExecutor.toExcelCellValue(
                new CellInput.DateTime(LocalDateTime.of(2026, 3, 23, 10, 15, 30))));
    assertEquals(LocalDateTime.of(2026, 3, 23, 10, 15, 30), dateTimeValue.value());
  }

  @Test
  void validatesTypedInputRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Text(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Numeric(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.BooleanValue(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Formula(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Formula("="));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Formula("=   "));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Date(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.DateTime(null));
  }
}
