package dev.erst.gridgrind.core.vocabulary;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Independent A1 boundary examples for the authored coordinate contract. */
class ProtocolCellAddressValidationTest {
  @Test
  void acceptsTheLastXlsxCoordinateAndRejectsTheNextRowAndColumn() {
    assertEquals("XFD1048576", ProtocolCellAddressValidation.validateAddress("XFD1048576"));
    assertEquals("$xfd$1048576", ProtocolCellAddressValidation.validateAddress("$xfd$1048576"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolCellAddressValidation.validateAddress("XFE1"));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolCellAddressValidation.validateAddress("A1048577"));
  }
}
