package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertIterableEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests selector JSON type-id registry behavior for canonical root-level serialization. */
class SelectorJsonSupportTest {
  @Test
  void resolvesKnownSelectorLeafTypeIds() {
    assertEquals("WORKBOOK_CURRENT", SelectorJsonSupport.typeIdFor(WorkbookSelector.Current.class));
    assertEquals("CELL_BY_ADDRESS", SelectorJsonSupport.typeIdFor(CellSelector.ByAddress.class));
    assertEquals(
        "NAMED_RANGE_WORKBOOK_SCOPE",
        SelectorJsonSupport.typeIdFor(NamedRangeSelector.WorkbookScope.class));
    assertTrue(
        SelectorJsonSupport.supportsTypeId(
            NamedRangeSelector.WorkbookScope.class, "NAMED_RANGE_WORKBOOK_SCOPE"));
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

    IllegalStateException duplicateTypeId =
        assertThrows(
            IllegalStateException.class,
            () ->
                SelectorJsonSupport.buildTypeIds(
                    List.of(DuplicateRootOne.class, DuplicateRootTwo.class)));
    assertEquals(
        "Duplicate selector type id 'DUPLICATE' for "
            + DuplicateSelectorAlpha.class.getName()
            + " and "
            + DuplicateSelectorBeta.class.getName(),
        duplicateTypeId.getMessage());
  }

  @Test
  void allowsRepeatedClaimsWhenTheWireIdStillPointsAtTheSameRuntimeType() {
    assertEquals(
        "REUSED",
        SelectorJsonSupport.buildTypeIds(List.of(RepeatedRootOne.class, RepeatedRootTwo.class))
            .get(RepeatedSelector.class));
  }

  @Test
  void familyInfoNormalizesValuesAndRejectsInvalidShapes() {
    SelectorJsonSupport.FamilyInfo familyInfo =
        new SelectorJsonSupport.FamilyInfo(
            "TableSelector", List.of("TABLE_BY_NAME", "TABLE_BY_NAME"));

    assertEquals("TableSelector", familyInfo.family());
    assertIterableEquals(List.of("TABLE_BY_NAME", "TABLE_BY_NAME"), familyInfo.typeIds());
    assertEquals(
        "family must not be null",
        assertThrows(
                NullPointerException.class,
                () -> new SelectorJsonSupport.FamilyInfo(null, List.of("TABLE_BY_NAME")))
            .getMessage());
    assertEquals(
        "family must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SelectorJsonSupport.FamilyInfo(" ", List.of("TABLE_BY_NAME")))
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
                        "TableSelector", java.util.Arrays.asList("TABLE_BY_NAME", null)))
            .getMessage());
    assertEquals(
        "typeIds must not contain blank values",
        assertThrows(
                IllegalArgumentException.class,
                () -> new SelectorJsonSupport.FamilyInfo("TableSelector", List.of(" ")))
            .getMessage());
  }

  /** Synthetic selector root used to prove duplicate type-id detection. */
  @JsonSubTypes(@JsonSubTypes.Type(value = DuplicateSelectorAlpha.class, name = "DUPLICATE"))
  private interface DuplicateRootOne {}

  /** First synthetic selector leaf that claims the duplicated wire id. */
  private static final class DuplicateSelectorAlpha {}

  /** Second synthetic selector root used to trigger the duplicate-wire-id guard. */
  @JsonSubTypes(@JsonSubTypes.Type(value = DuplicateSelectorBeta.class, name = "DUPLICATE"))
  private interface DuplicateRootTwo {}

  /** Second synthetic selector leaf that collides on the same wire id. */
  private static final class DuplicateSelectorBeta {}

  /** Synthetic selector root that reuses a wire id for the same runtime type. */
  @JsonSubTypes(@JsonSubTypes.Type(value = RepeatedSelector.class, name = "REUSED"))
  private interface RepeatedRootOne {}

  /** Second synthetic selector root that repeats the same runtime-type ownership. */
  @JsonSubTypes(@JsonSubTypes.Type(value = RepeatedSelector.class, name = "REUSED"))
  private interface RepeatedRootTwo {}

  /** Synthetic selector leaf reused across multiple roots without creating ambiguity. */
  private static final class RepeatedSelector {}
}
