package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.*;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage tests for standalone field-shape value objects. */
class FieldShapeTest {
  @Test
  void scalarListAndNestedTypeGroupsValidateRequiredValues() {
    assertEquals(ScalarType.STRING, new FieldShape.Scalar(ScalarType.STRING).scalarType());
    assertEquals(
        new FieldShape.Scalar(ScalarType.STRING),
        new FieldShape.ListShape(new FieldShape.Scalar(ScalarType.STRING)).elementShape());
    assertEquals(
        "sheetSelectorTypes", new FieldShape.NestedTypeGroupRef("sheetSelectorTypes").group());
    assertEquals("commentInputType", new FieldShape.PlainTypeGroupRef("commentInputType").group());

    assertThrows(NullPointerException.class, () -> new FieldShape.Scalar(null));
    assertThrows(NullPointerException.class, () -> new FieldShape.ListShape(null));
    assertThrows(NullPointerException.class, () -> new FieldShape.NestedTypeGroupRef(null));
    assertThrows(IllegalArgumentException.class, () -> new FieldShape.NestedTypeGroupRef(" "));
    assertThrows(NullPointerException.class, () -> new FieldShape.PlainTypeGroupRef(null));
    assertThrows(IllegalArgumentException.class, () -> new FieldShape.PlainTypeGroupRef(" "));
  }

  @Test
  void topLevelTypeSetRefRequiresNonBlankTypeSet() {
    assertEquals("operationTypes", new FieldShape.TopLevelTypeSetRef("operationTypes").typeSet());
    assertThrows(NullPointerException.class, () -> new FieldShape.TopLevelTypeSetRef(null));
    assertThrows(IllegalArgumentException.class, () -> new FieldShape.TopLevelTypeSetRef(" "));
  }

  @Test
  void nestedTypeGroupUnionRequiresNonBlankGroups() {
    assertEquals(
        List.of("sheetSelectorTypes", "cellSelectorTypes"),
        new FieldShape.NestedTypeGroupUnionRef(List.of("sheetSelectorTypes", "cellSelectorTypes"))
            .groups());
    assertThrows(NullPointerException.class, () -> new FieldShape.NestedTypeGroupUnionRef(null));
    assertThrows(
        IllegalArgumentException.class, () -> new FieldShape.NestedTypeGroupUnionRef(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new FieldShape.NestedTypeGroupUnionRef(List.of("sheetSelectorTypes", " ")));
  }
}
