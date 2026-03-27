package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for cell-selection request invariants. */
class CellSelectionTest {
  @Test
  void selectedRejectsBlankAddresses() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () -> new CellSelection.Selected(Arrays.asList("A1", " ")));

    assertEquals("addresses must not contain blank values", exception.getMessage());
  }

  @Test
  void selectedCopiesAddresses() {
    List<String> addresses = Arrays.asList("A1", "B2");
    CellSelection.Selected selection = new CellSelection.Selected(addresses);

    assertEquals(List.of("A1", "B2"), selection.addresses());
  }
}
