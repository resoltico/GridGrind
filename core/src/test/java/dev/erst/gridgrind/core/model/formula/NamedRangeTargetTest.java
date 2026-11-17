package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeTarget authored construction and validation. */
class NamedRangeTargetTest {
  @Test
  void usesCanonicalOrderedRangeGeometry() {
    NamedRangeTarget.Range target =
        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("$a$1:B4"));

    assertEquals("A1:B4", target.range().value());
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4:A1")));
    assertEquals("Budget!$B$4:A1", NamedRangeTarget.formula("Budget!$B$4:A1").formula());
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(
        NullPointerException.class, () -> NamedRangeTarget.range(null, new RangeRef("A1")));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range(new SheetName(" "), new RangeRef("A1")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            NamedRangeTarget.range(
                new SheetName("ThisSheetNameIsFarTooLongForExcel"), new RangeRef("A1")));
    assertThrows(
        NullPointerException.class, () -> NamedRangeTarget.range(new SheetName("Budget"), null));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.range(new SheetName("Budget"), new RangeRef(" ")));
  }
}
