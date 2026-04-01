package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for ExcelRangeSelection constructor invariants and defensive copies. */
class ExcelRangeSelectionTest {
  @Test
  void createsAllSelection() {
    ExcelRangeSelection.All selection = new ExcelRangeSelection.All();

    assertInstanceOf(ExcelRangeSelection.All.class, selection);
  }

  @Test
  void selectedCopiesRangesAndPreservesOrder() {
    List<String> ranges = new ArrayList<>(List.of("A1:A3", "C2:D4"));

    ExcelRangeSelection.Selected selection = new ExcelRangeSelection.Selected(ranges);
    ranges.clear();

    assertEquals(List.of("A1:A3", "C2:D4"), selection.ranges());
  }

  @Test
  void selectedRejectsInvalidRangeCollections() {
    assertThrows(NullPointerException.class, () -> new ExcelRangeSelection.Selected(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRangeSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelRangeSelection.Selected(List.of("A1:A3", "A1:A3")));
    assertThrows(
        NullPointerException.class, () -> new ExcelRangeSelection.Selected(List.of("A1:A3", null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelRangeSelection.Selected(List.of("A1:A3", " ")));
  }
}
