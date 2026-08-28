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
            ExcelNamedRangeTarget.range("Budget", "B4"));

    assertEquals("BudgetTotal", definition.name());
    assertEquals("Ieņēmumi", ExcelNamedRangeDefinition.validateName("Ieņēmumi"));
    assertEquals("Доходы", ExcelNamedRangeDefinition.validateName("Доходы"));
    assertEquals("収益", ExcelNamedRangeDefinition.validateName("収益"));
    assertEquals("\\Ledger", ExcelNamedRangeDefinition.validateName("\\Ledger"));
    assertEquals("Ledger\\2026", ExcelNamedRangeDefinition.validateName("Ledger\\2026"));
    assertEquals("Ledger.2026", ExcelNamedRangeDefinition.validateName("Ledger.2026"));
    assertEquals("LedgerⅫ", ExcelNamedRangeDefinition.validateName("LedgerⅫ"));
    assertEquals("Ledger¼", ExcelNamedRangeDefinition.validateName("Ledger¼"));
    assertEquals("𐐀Ledger", ExcelNamedRangeDefinition.validateName("𐐀Ledger"));
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
    assertEquals("XFE1", ExcelNamedRangeDefinition.validateName("XFE1"));
    assertEquals("𐐀".repeat(255), ExcelNamedRangeDefinition.validateName("𐐀".repeat(255)));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelNamedRangeDefinition.validateName("𐐀".repeat(256)));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeDefinition(
                "BudgetTotal", null, ExcelNamedRangeTarget.range("Budget", "B4")));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeDefinition(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope(), null));
  }

  @Test
  void observedNamesRequireOnlyNonblankFactualText() {
    assertEquals("9legacy name", ExcelNamedRangeDefinition.validateObservedName("9legacy name"));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelNamedRangeDefinition.validateObservedName(" "));
  }
}
