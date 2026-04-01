package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for range-selection record construction. */
class RangeSelectionTest {
  @Test
  void selectedCopiesRanges() {
    List<String> ranges = new ArrayList<>(List.of("A1:B2", "D4"));

    RangeSelection.Selected selection = new RangeSelection.Selected(ranges);
    ranges.clear();

    assertEquals(List.of("A1:B2", "D4"), selection.ranges());
  }

  @Test
  void validatesSelectionInputs() {
    assertInstanceOf(RangeSelection.All.class, new RangeSelection.All());
    assertThrows(NullPointerException.class, () -> new RangeSelection.Selected(null));
    assertThrows(IllegalArgumentException.class, () -> new RangeSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new RangeSelection.Selected(List.of("A1", " ")));
    assertThrows(
        IllegalArgumentException.class, () -> new RangeSelection.Selected(List.of("A1", "A1")));
  }
}
