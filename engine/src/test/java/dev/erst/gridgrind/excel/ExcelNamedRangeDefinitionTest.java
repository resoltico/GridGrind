package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeDefinition validation rules. */
class ExcelNamedRangeDefinitionTest {
  @Test
  void acceptsValidNamedRangeDefinitions() {
    ExcelNamedRangeDefinition definition =
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("Budget", "B4"));

    assertEquals("BudgetTotal", definition.name());
  }

  @Test
  void rejectsInvalidNamedRangeNamesAndNullComponents() {
    assertThrows(NullPointerException.class, () -> ExcelNamedRangeDefinition.validateName(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelNamedRangeDefinition.validateName(" "));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelNamedRangeDefinition.validateName("1Budget"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelNamedRangeDefinition.validateName("_xlnm.Print_Area"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelNamedRangeDefinition.validateName("_XLNM.PRINT_AREA"));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelNamedRangeDefinition.validateName("A1"));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelNamedRangeDefinition.validateName("R1C1"));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeDefinition(
                "BudgetTotal", null, new ExcelNamedRangeTarget("Budget", "B4")));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeDefinition(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope(), null));
  }
}
