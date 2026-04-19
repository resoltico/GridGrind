package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelRichText;
import dev.erst.gridgrind.excel.ExcelRichTextRun;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for CellInput record construction and ExcelCellValue conversion. */
class CellInputTest {
  @Test
  void convertsAllSupportedInputTypesToExcelValues() {
    assertInstanceOf(
        ExcelCellValue.BlankValue.class,
        WorkbookCommandConverter.toExcelCellValue(new CellInput.Blank()));

    ExcelCellValue.TextValue textValue =
        assertInstanceOf(
            ExcelCellValue.TextValue.class,
            WorkbookCommandConverter.toExcelCellValue(textCell("Budget")));
    assertEquals("Budget", textValue.value());

    ExcelCellValue.RichTextValue richTextValue =
        assertInstanceOf(
            ExcelCellValue.RichTextValue.class,
            WorkbookCommandConverter.toExcelCellValue(
                new CellInput.RichText(
                    List.of(
                        richTextRun("Budget"),
                        richTextRun(
                            " FY26",
                            new CellFontInput(
                                Boolean.TRUE, null, null, null, "#AABBCC", null, null))))));
    assertEquals(
        new ExcelRichText(
            List.of(
                new ExcelRichTextRun("Budget", null),
                new ExcelRichTextRun(
                    " FY26",
                    new dev.erst.gridgrind.excel.ExcelCellFont(
                        Boolean.TRUE,
                        null,
                        null,
                        null,
                        new dev.erst.gridgrind.excel.ExcelColor("#AABBCC"),
                        null,
                        null)))),
        richTextValue.value());

    ExcelCellValue.NumberValue numberValue =
        assertInstanceOf(
            ExcelCellValue.NumberValue.class,
            WorkbookCommandConverter.toExcelCellValue(new CellInput.Numeric(42.5)));
    assertEquals(42.5, numberValue.value());

    ExcelCellValue.BooleanValue booleanValue =
        assertInstanceOf(
            ExcelCellValue.BooleanValue.class,
            WorkbookCommandConverter.toExcelCellValue(new CellInput.BooleanValue(true)));
    assertTrue(booleanValue.value());

    ExcelCellValue.FormulaValue formulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            WorkbookCommandConverter.toExcelCellValue(formulaCell("SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", formulaValue.expression());

    // Leading = is stripped automatically so callers can use Excel-native syntax
    ExcelCellValue.FormulaValue strippedFormulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            WorkbookCommandConverter.toExcelCellValue(formulaCell("=SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", strippedFormulaValue.expression());

    ExcelCellValue.DateValue dateValue =
        assertInstanceOf(
            ExcelCellValue.DateValue.class,
            WorkbookCommandConverter.toExcelCellValue(
                new CellInput.Date(LocalDate.of(2026, 3, 23))));
    assertEquals(LocalDate.of(2026, 3, 23), dateValue.value());

    ExcelCellValue.DateTimeValue dateTimeValue =
        assertInstanceOf(
            ExcelCellValue.DateTimeValue.class,
            WorkbookCommandConverter.toExcelCellValue(
                new CellInput.DateTime(LocalDateTime.of(2026, 3, 23, 10, 15, 30))));
    assertEquals(LocalDateTime.of(2026, 3, 23, 10, 15, 30), dateTimeValue.value());
  }

  @Test
  void validatesTypedInputRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Text(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.RichText(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.RichText(List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new CellInput.RichText(List.of(richTextRun(""))));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Numeric(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.BooleanValue(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Formula(null));
    assertThrows(IllegalArgumentException.class, () -> formulaCell("="));
    assertThrows(IllegalArgumentException.class, () -> formulaCell("=   "));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Date(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.DateTime(null));
  }
}
