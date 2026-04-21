package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertIterableEquals;
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

  @Test
  void familyInfoNormalizesValuesAndRejectsInvalidShapes() {
    SelectorJsonSupport.FamilyInfo familyInfo =
        new SelectorJsonSupport.FamilyInfo("TableSelector", List.of("BY_NAME", "BY_NAME"));

    assertEquals("TableSelector", familyInfo.family());
    assertIterableEquals(List.of("BY_NAME", "BY_NAME"), familyInfo.typeIds());
    assertEquals(
        "family must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new SelectorJsonSupport.FamilyInfo(null, List.of("BY_NAME")))
            .getMessage());
    assertEquals(
        "family must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SelectorJsonSupport.FamilyInfo(" ", List.of("BY_NAME")))
            .getMessage());
    assertEquals(
        "typeIds must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new SelectorJsonSupport.FamilyInfo("TableSelector", null))
            .getMessage());
    assertEquals(
        "typeIds must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SelectorJsonSupport.FamilyInfo("TableSelector", List.of()))
            .getMessage());
    assertEquals(
        "typeIds must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new SelectorJsonSupport.FamilyInfo(
                        "TableSelector", java.util.Arrays.asList("BY_NAME", null)))
            .getMessage());
    assertEquals(
        "typeIds must not contain blank values",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SelectorJsonSupport.FamilyInfo("TableSelector", List.of(" ")))
            .getMessage());
  }
}
