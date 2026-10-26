package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.TableColumnSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ExcelTableColumnDefinition;
import dev.erst.gridgrind.excel.ExcelTableDefinition;
import dev.erst.gridgrind.excel.ExcelTableStyle;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table authoring input invariants and conversion. */
class TableSpecTest {
  @Test
  void validatesAndConvertsToWorkbookDefinition() {
    TableSpec input =
        TableSpec.withDefaultMetadata(
            new TableName("BudgetTable"),
            new SheetName("Budget"),
            new RangeRef("A1:C4"),
            true,
            new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false));

    assertEquals(
        new ExcelTableDefinition(
            "BudgetTable",
            "Budget",
            "A1:C4",
            true,
            new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false)),
        WorkbookCommandTabularInputConverter.toExcelTableDefinition(input));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            TableSpec.withDefaultMetadata(
                new TableName("BudgetTable"),
                new SheetName(" "),
                new RangeRef("A1:C4"),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            TableSpec.withDefaultMetadata(
                new TableName("BudgetTable"),
                new SheetName("12345678901234567890123456789012"),
                new RangeRef("A1:C4"),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            TableSpec.withDefaultMetadata(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef(" "),
                false,
                new TableStyleSpec.None()));
    assertThrows(
        NullPointerException.class,
        () ->
            TableSpec.withDefaultMetadata(
                new TableName("BudgetTable"),
                new SheetName("Budget"),
                new RangeRef("A1:C4"),
                false,
                null));
  }

  @Test
  void explicitDefaultMetadataPreservesTotalsRowFlagFalse() {
    TableSpec input =
        TableSpec.withDefaultMetadata(
            new TableName("BudgetTable"),
            new SheetName("Budget"),
            new RangeRef("A1:C4"),
            false,
            new TableStyleSpec.None());

    assertFalse(input.showTotalsRow());
    assertEquals(
        new ExcelTableDefinition(
            "BudgetTable", "Budget", "A1:C4", false, new ExcelTableStyle.None()),
        WorkbookCommandTabularInputConverter.toExcelTableDefinition(input));
  }

  @Test
  void normalizesTotalsRowFunctionsToLowercaseForProtocolWrites() {
    TableSpec input =
        new TableSpec(
            new TableName("BudgetTable"),
            new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
            new RangeRef("A1:C4"),
            true,
            true,
            new TableStyleSpec.None(),
            dev.erst.gridgrind.core.source.TextSourceSpec.inline(""),
            false,
            false,
            false,
            "",
            "",
            "",
            List.of(new TableColumnSpec(1, "", "Total", " SUM ", "")));

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
        WorkbookCommandTabularInputConverter.toExcelTableDefinition(input));
  }
}
