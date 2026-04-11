package dev.erst.gridgrind.protocol.catalog;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Coverage tests for standalone field-shape value objects. */
class FieldShapeTest {
  @Test
  void topLevelTypeSetRefRequiresNonBlankTypeSet() {
    assertEquals("operationTypes", new FieldShape.TopLevelTypeSetRef("operationTypes").typeSet());
    assertThrows(NullPointerException.class, () -> new FieldShape.TopLevelTypeSetRef(null));
    assertThrows(IllegalArgumentException.class, () -> new FieldShape.TopLevelTypeSetRef(" "));
  }
}
