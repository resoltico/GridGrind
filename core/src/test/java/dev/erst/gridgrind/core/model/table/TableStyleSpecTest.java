package dev.erst.gridgrind.core.model.table;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Checks named table-style requirements. */
class TableStyleSpecTest {
  @Test
  void validatesNamedStyleIdentifiers() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableStyleSpec.Named(" ", false, false, true, false));
    assertThrows(
        NullPointerException.class,
        () -> new TableStyleSpec.Named(null, false, false, true, false));
  }
}
