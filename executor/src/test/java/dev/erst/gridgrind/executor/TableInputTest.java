package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.TableColumnInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.excel.ExcelTableColumnDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table authoring input invariants and conversion. */
class TableInputTest {
  @Test
  void validatesAndConvertsToWorkbookDefinition() {
    TableInput input =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            true,
            new TableStyleInput.Named("TableStyleMedium2", false, false, true, false));

    assertEquals(
        new ExcelTableDefinition(
            "BudgetTable",
            "Budget",
            "A1:C4",
            true,
            new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false)),
        WorkbookCommandConverter.toExcelTableDefinition(input));

    assertThrows(
        IllegalArgumentException.class,
        () -> new TableInput("BudgetTable", " ", "A1:C4", false, new TableStyleInput.None()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableInput(
                "BudgetTable",
                "12345678901234567890123456789012",
                "A1:C4",
                false,
                new TableStyleInput.None()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableInput("BudgetTable", "Budget", " ", false, new TableStyleInput.None()));
    assertThrows(
        NullPointerException.class,
        () -> new TableInput("BudgetTable", "Budget", "A1:C4", false, null));
  }

  @Test
  void defaultsMissingTotalsRowFlagToFalse() {
    TableInput input =
        new TableInput("BudgetTable", "Budget", "A1:C4", null, new TableStyleInput.None());

    assertFalse(input.showTotalsRow());
    assertEquals(
        new ExcelTableDefinition(
            "BudgetTable", "Budget", "A1:C4", false, new ExcelTableStyle.None()),
        WorkbookCommandConverter.toExcelTableDefinition(input));
  }

  @Test
  void normalizesTotalsRowFunctionsToLowercaseForProtocolWrites() {
    TableInput input =
        new TableInput(
            "BudgetTable",
            "Budget",
            "A1:C4",
            true,
            true,
            new TableStyleInput.None(),
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            List.of(new TableColumnInput(1, "", "Total", " SUM ", "")));

    assertEquals(
        new ExcelTableDefinition(
            "BudgetTable",
            "Budget",
            "A1:C4",
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
            List.of(new ExcelTableColumnDefinition(1, "", "Total", "sum", ""))),
        WorkbookCommandConverter.toExcelTableDefinition(input));
  }
}
