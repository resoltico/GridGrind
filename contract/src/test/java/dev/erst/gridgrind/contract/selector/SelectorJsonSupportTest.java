package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests selector JSON type-id registry behavior for canonical root-level serialization. */
class SelectorJsonSupportTest {
  @Test
  void resolvesKnownSelectorLeafTypeIds() {
    assertEquals("CURRENT", SelectorJsonSupport.typeIdFor(WorkbookSelector.Current.class));
    assertEquals("BY_ADDRESS", SelectorJsonSupport.typeIdFor(CellSelector.ByAddress.class));
    assertEquals(
        "WORKBOOK_SCOPE", SelectorJsonSupport.typeIdFor(NamedRangeSelector.WorkbookScope.class));
  }

  @Test
  void rejectsUnsupportedRuntimeTypesAndRootsWithoutJsonSubtypes() {
    IllegalArgumentException unsupportedType =
        assertThrows(
            IllegalArgumentException.class, () -> SelectorJsonSupport.typeIdFor(String.class));
    assertEquals(
        "Unsupported selector runtime type: class java.lang.String", unsupportedType.getMessage());

    IllegalStateException missingSubtypes =
        assertThrows(
            IllegalStateException.class,
            () -> SelectorJsonSupport.buildTypeIds(List.of(String.class)));
    assertEquals(
        "Selector root is missing @JsonSubTypes: class java.lang.String",
        missingSubtypes.getMessage());
  }
}
