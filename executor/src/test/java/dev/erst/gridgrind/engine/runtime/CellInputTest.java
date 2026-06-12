package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelColor;
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
        WorkbookCommandCellInputConverter.toExcelCellValue(new CellInput.Blank()));

    ExcelCellValue.TextValue textValue =
        assertInstanceOf(
            ExcelCellValue.TextValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(textCell("Budget")));
    assertEquals("Budget", textValue.value());

    ExcelCellValue.RichTextValue richTextValue =
        assertInstanceOf(
            ExcelCellValue.RichTextValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(
                new CellInput.RichText(
                    List.of(
                        richTextRun("Budget"),
                        richTextRun(
                            " FY26",
                            fontInput(
                                Boolean.TRUE,
                                null,
                                null,
                                null,
                                ColorInput.rgb("#AABBCC"),
                                null,
                                null))))));
    assertEquals(
        new ExcelRichText(
            List.of(
                new ExcelRichTextRun("Budget", java.util.Optional.empty()),
                new ExcelRichTextRun(
                    " FY26",
                    java.util.Optional.of(
                        new dev.erst.gridgrind.excel.ExcelCellFont(
                            java.util.Optional.of(Boolean.TRUE),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.of(ExcelColor.rgb("#AABBCC")),
                            java.util.Optional.empty(),
                            java.util.Optional.empty()))))),
        richTextValue.value());

    ExcelCellValue.NumberValue numberValue =
        assertInstanceOf(
            ExcelCellValue.NumberValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(new CellInput.NumberValue(42.5)));
    assertEquals(42.5, numberValue.value());

    ExcelCellValue.BooleanValue booleanValue =
        assertInstanceOf(
            ExcelCellValue.BooleanValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(new CellInput.BooleanValue(true)));
    assertTrue(booleanValue.value());

    ExcelCellValue.FormulaValue formulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(formulaCell("SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", formulaValue.expression());

    // Leading = is stripped automatically so callers can use Excel-native syntax
    ExcelCellValue.FormulaValue strippedFormulaValue =
        assertInstanceOf(
            ExcelCellValue.FormulaValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(formulaCell("=SUM(B2:B4)")));
    assertEquals("SUM(B2:B4)", strippedFormulaValue.expression());

    ExcelCellValue.DateValue dateValue =
        assertInstanceOf(
            ExcelCellValue.DateValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(
                new CellInput.Date(LocalDate.of(2026, 3, 23))));
    assertEquals(LocalDate.of(2026, 3, 23), dateValue.value());

    ExcelCellValue.DateTimeValue dateTimeValue =
        assertInstanceOf(
            ExcelCellValue.DateTimeValue.class,
            WorkbookCommandCellInputConverter.toExcelCellValue(
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
    assertThrows(
        IllegalArgumentException.class, () -> new CellInput.NumberValue(Double.POSITIVE_INFINITY));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.NumberValue(Double.NaN));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Formula(null));
    assertThrows(IllegalArgumentException.class, () -> formulaCell("="));
    assertThrows(IllegalArgumentException.class, () -> formulaCell("=   "));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.Date(null));
    assertThrows(IllegalArgumentException.class, () -> new CellInput.DateTime(null));
  }

  @Test
  void rejectsDdeFormulaInjectionInAllCaseForms() {
    IllegalArgumentException ddeUpper =
        assertThrows(
            IllegalArgumentException.class, () -> formulaCell("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertTrue(ddeUpper.getMessage().contains("DDE"));

    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("Dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("=DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class, () -> formulaCell("=dde(\"cmd\",\"/C calc\",\"\")"));
  }
}
