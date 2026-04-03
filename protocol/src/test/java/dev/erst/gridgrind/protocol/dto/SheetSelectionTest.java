package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for sheet-selection request invariants. */
class SheetSelectionTest {
  @Test
  void selectedRejectsBlankSheetNames() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> new SheetSelection.Selected(Arrays.asList("Budget", " ")));

    assertEquals("sheetNames must not contain blank values", exception.getMessage());
  }

  @Test
  void selectedCopiesSheetNames() {
    List<String> sheetNames = Arrays.asList("Budget", "Summary");
    SheetSelection.Selected selection = new SheetSelection.Selected(sheetNames);

    assertEquals(List.of("Budget", "Summary"), selection.sheetNames());
  }
}
