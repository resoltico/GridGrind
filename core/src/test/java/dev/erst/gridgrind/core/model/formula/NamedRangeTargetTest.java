package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeTarget authored construction and validation. */
class NamedRangeTargetTest {
  @Test
  void preservesAuthoredRangeText() {
    NamedRangeTarget.Range target =
        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4:A1"));

    assertEquals("B4:A1", target.range().value());
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
