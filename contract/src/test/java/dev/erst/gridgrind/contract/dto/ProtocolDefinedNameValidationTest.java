package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for protocol-owned defined-name validation helpers. */
class ProtocolDefinedNameValidationTest {
  @Test
  void acceptsValidNamesWithinTheShippedContract() {
    assertEquals("Budget_Total", ProtocolDefinedNameValidation.validateName("Budget_Total"));
    assertEquals("XFE1", ProtocolDefinedNameValidation.validateName("XFE1"));
  }

  @Test
  void rejectsNamesThatBreakWorkbookIdentifierRules() {
    assertThrows(
        NullPointerException.class, () -> ProtocolDefinedNameValidation.validateName(null));
    assertThrows(
        IllegalArgumentException.class, () -> ProtocolDefinedNameValidation.validateName(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolDefinedNameValidation.validateName("9Budget"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolDefinedNameValidation.validateName("Budget Total"));
  }

  @Test
  void rejectsReservedAndReferenceLikeNames() {
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolDefinedNameValidation.validateName("_xlnm.Print_Area"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolDefinedNameValidation.validateName("_XLNM.Print_Area"));
    assertThrows(
        IllegalArgumentException.class, () -> ProtocolDefinedNameValidation.validateName("A1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolDefinedNameValidation.validateName("XFD1048576"));
    assertThrows(
        IllegalArgumentException.class, () -> ProtocolDefinedNameValidation.validateName("R12C3"));
  }
}
