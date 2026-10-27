package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import org.junit.jupiter.api.Test;

/** Tests for StructuredMutationAction.SetNamedRange validation rules. */
class NamedRangeActionValidationTest {
  @Test
  void acceptsValidNamedRangeDefinitions() {
    StructuredMutationAction.SetNamedRange definition =
        new StructuredMutationAction.SetNamedRange(
            new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
            new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")));

    assertEquals("BudgetTotal", definition.name().value());
    assertEquals(
        "Ieņēmumi",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("Ieņēmumi"));
    assertEquals(
        "Доходы",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("Доходы"));
    assertEquals(
        "収益", dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("収益"));
    assertEquals(
        "\\Ledger",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("\\Ledger"));
    assertEquals(
        "Ledger\\2026",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
            "Ledger\\2026"));
    assertEquals(
        "Ledger.2026",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
            "Ledger.2026"));
    assertEquals(
        "LedgerⅫ",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("LedgerⅫ"));
    assertEquals(
        "Ledger¼",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("Ledger¼"));
    assertEquals(
        "𐐀Ledger",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("𐐀Ledger"));
  }

  @Test
  void rejectsInvalidNamedRangeNamesAndNullComponents() {
    assertThrows(
        NullPointerException.class,
        () -> dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(" "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
                "1Budget"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
                "_xlnm.Print_Area"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
                "_XLNM.PRINT_AREA"));
    assertThrows(
        IllegalArgumentException.class,
        () -> dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("A1"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("R1C1"));
    assertEquals(
        "XFE1",
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName("XFE1"));
    assertEquals(
        "𐐀".repeat(255),
        dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
            "𐐀".repeat(255)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation.validateName(
                "𐐀".repeat(256)));
    assertThrows(
        NullPointerException.class,
        () ->
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                null,
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B4"))));
    assertThrows(
        NullPointerException.class,
        () ->
            new StructuredMutationAction.SetNamedRange(
                new dev.erst.gridgrind.core.vocabulary.DefinedName("BudgetTotal"),
                new dev.erst.gridgrind.core.vocabulary.NamedRangeScope.Workbook(),
                null));
  }

  @Test
  void observedNamesRequireOnlyNonblankFactualText() {
    assertEquals(
        "9legacy name",
        new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("9legacy name").value());
    assertThrows(
        IllegalArgumentException.class,
        () -> new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel(" "));
  }
}
